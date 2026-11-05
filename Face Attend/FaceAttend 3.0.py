"""
FaceAttend 3.0 ·  Smart Attendance System (Mega Feature Upgrade)
======================================================
Author  : Built on FaceAttend 2.0 by the same team
Version : 3.0
Build   : 2026.08

What's new in 3.0:
--------------------------
• Multi-select roster table: Shift-click selects a range, Ctrl-click toggles
  individual rows (native Treeview "extended" select), plus a one-click
  "Select All" button. Remove and section-change actions now operate on
  every selected student at once instead of one at a time.
• Year-end Promotion tool (Settings → Tools → "Promote Students"): reads
  every distinct grade/section in the roster (e.g. 7C), auto-suggests the
  next grade for each (7C → 8C) while keeping the section letter, and lets
  the teacher override any individual mapping (e.g. reshuffle 7C → 8B) in
  one table before applying it to the whole school in a single bulk write —
  no more editing students one-by-one every year.
• AI Buddy is now an actual actuator, not just a Q&A bot. It understands
  bulk natural-language commands such as "remove these students - Aditi,
  Rohan, Meera" or "change these students to section A: Kabir, Zara" and
  carries them out (with a confirmation prompt) against the real roster,
  in addition to everything it already answered before.
• Two new tabs: Whiteboard (freehand drawing canvas with color/size picker,
  saved as PNG) and Notepad (plain-text notes editor, saved as .txt) — both
  save into dedicated folders inside the FaceAttend data directory so
  a teacher's scratch notes and diagrams live alongside the roster.
• Language selector actually works now: 13 distinct visual themes — one
  per supported language — each with its own accent palette, font pairing
  and rail styling, applied on restart so switching language genuinely
  changes how the app looks, not just a handful of translated labels.
• Rebuilt startup screen: package/module checklist now sizes itself to fit
  every row so nothing scrolls out of view, has a real determinate
  progress bar, and a watchdog timer so the app always reaches the main
  window instead of appearing to hang.

Carried over from 2.10 (packaged-build crash fix):
--------------------------
• Fixed a hard crash (Windows APPCRASH, access violation in MSVCP140.dll)
  that only happened in the packaged .exe build, not when running from
  source. Root cause: the bundled `dlib` wheel used CPU instructions
  (AVX) not supported on some machines, causing the face-recognition
  engine to crash the whole app instead of failing gracefully. Fixed by
  switching to the `dlib-bin` package, which ships a more broadly
  CPU-compatible prebuilt binary.
• face_recognition's import failure is no longer silently swallowed —
  the real exception (if any) is now written to
  face_recognition_error.log next to the app, so future packaging
  issues are diagnosable without guesswork.

Carried over from 2.9 (2026-06 stability patch):
--------------------------
• Fixed a hard crash (Windows APPCRASH, access violation in MSVCP140.dll)
  that only happened in the packaged .exe build, not when running from
  source. Root cause: the bundled `dlib` wheel used CPU instructions
  (AVX) not supported on some machines, causing the face-recognition
  engine to crash the whole app instead of failing gracefully. Fixed by
  switching to the `dlib-bin` package, which ships a more broadly
  CPU-compatible prebuilt binary.
• face_recognition's import failure is no longer silently swallowed —
  the real exception (if any) is now written to
  face_recognition_error.log next to the app, so future packaging
  issues are diagnosable without guesswork.

Carried over from 2.9 (2026-06 stability patch):
--------------------------
• Single Tk root — splash screen now uses a Toplevel instead of spawning a
  second root window, eliminating the classic tkinter multi-root crash risk.
• Audit log CSV header now matches the 10 data columns (added House, Streak).
• Settings Reset now actually persists the reset values to disk.
• Auto-notify email/WhatsApp decoupled from UI variable state — reads directly
  from Settings so the Mail tab no longer depends on Settings tab widgets.
• Camera loop, face-processing, and drawing are fully wrapped in try/except
  so one bad frame cannot kill the live scan.
• Tooltip, AnimBtn, Toast, and StatCard all defend against widget destruction
  mid-animation (TclError silenced gracefully).
• Unknown-face crop is resized if oversized to keep data/unknowns tidy.
• open_camera falls back to default backend if CAP_DSHOW fails on Windows.
• StudentDialog mousewheel binding is tagged and unbound cleanly on destroy.
• Excel image-row removal is wrapped for openpyxl version compatibility.
• WhatsApp Web first-open now checks webbrowser.open return value.
• Export menu is cached to avoid transient Menu GC issues.
• Activity-feed empty hint logic is more robust.
• Removed duplicate Tooltip on the rail status dot.
• Removed spurious pack_propagate on a grid-managed frame.
• Message body starts blank instead of containing placeholder text.
• Various defensive null checks and exception guards throughout.

Real feature set (for reference — not all of this was in 2.9's comments):
--------------------------
• Face-recognition attendance with liveness detection (hog-based
  face_recognition + dlib), plus QR-code attendance as a fallback path.
• Gamification layer: XP, streaks, badges, and house-vs-house competition.
• Matplotlib-based analytics dashboards.
• Ollama-backed local AI assistant ("AI Buddy") for querying attendance data.
• Parent notifications via Gmail and WhatsApp Web.
• UI text supports 13 Indian languages.
• Small built-in REST API (http.server-based) for external queries.
• Excel-native data store with self-healing/resilience if the workbook
  gets corrupted or a column shifts, so it doesn't misattribute attendance.
─────────────────────────────────────────────────────────────────
"""

from __future__ import annotations
import sys
import time

import csv
import io
import json
import logging
import os
import pickle
import re
import shutil
import subprocess
import threading
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from datetime import date, datetime, time as dtime
from pathlib import Path
from tkinter import messagebox, ttk, filedialog, colorchooser, simpledialog
import tkinter as tk

# ── version ─────────────────────────────────────────────────────────────────
APP_VERSION = "3.0"
APP_BUILD   = "2026.08"

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("faceattend")

# ── optional heavy packages ──────────────────────────────────────────────────
MISSING_PACKAGES: list[str] = []
face_recognition = None
FACE_RECOGNITION_IMPORT_ERROR = ""

try:
    import cv2
except Exception:
    cv2 = None
    MISSING_PACKAGES.append("opencv-python")
try:
    import numpy as np
except Exception:
    np = None
    MISSING_PACKAGES.append("numpy")
try:
    from PIL import Image, ImageTk, ImageDraw
except Exception:
    Image = ImageTk = ImageDraw = None
    MISSING_PACKAGES.append("Pillow")
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.drawing.image import Image as ExcelImage
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
except Exception:
    Workbook = load_workbook = ExcelImage = get_column_letter = None
    PatternFill = Font = Alignment = Border = Side = None
    MISSING_PACKAGES.append("openpyxl")
try:
    import pygame
    pygame.mixer.init()
    pygame.mixer.quit()
except Exception:
    pygame = None
    MISSING_PACKAGES.append("pygame")
try:
    import dlib
except Exception:
    dlib = None
try:
    import cmake  # noqa: F401 — only needed at build time for dlib
except Exception:
    pass
try:
    import face_recognition as _fr_mod
    face_recognition = _fr_mod
    try:
        import face_recognition_models  # noqa: F401
    except Exception:
        pass
except Exception as _fr_exc:
    face_recognition = None
    MISSING_PACKAGES.append("face_recognition")
    FACE_RECOGNITION_IMPORT_ERROR = f"{type(_fr_exc).__name__}: {_fr_exc}"
    try:
        _log_dir = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
        with open(_log_dir / "face_recognition_error.log", "w", encoding="utf-8") as _f:
            import traceback
            _f.write(FACE_RECOGNITION_IMPORT_ERROR + "\n\n")
            traceback.print_exc(file=_f)
    except Exception:
        pass
try:
    import pyttsx3
except Exception:
    pyttsx3 = None
try:
    from reportlab.lib import colors as _rl_colors  # noqa: F401
    HAS_REPORTLAB = True
except Exception:
    HAS_REPORTLAB = False
try:
    import qrcode
    from qrcode.image.pil import PilImage
except Exception:
    qrcode = None
    PilImage = None
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
except Exception:
    plt = Figure = FigureCanvasTkAgg = None
try:
    from cryptography.fernet import Fernet
except Exception:
    Fernet = None
try:
    from scipy.spatial import distance as sci_dist
except Exception:
    sci_dist = None
try:
    from http.server import BaseHTTPRequestHandler, HTTPServer
    import socketserver
except Exception:
    BaseHTTPRequestHandler = HTTPServer = socketserver = None

# ── paths ────────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

DATA_DIR        = APP_DIR / "data"
FACES_DIR       = DATA_DIR / "faces"
UNKNOWNS_DIR    = DATA_DIR / "unknowns"
ASSETS_DIR      = APP_DIR / "assets"
ICON_FILE       = ASSETS_DIR / "faceattend_icon.ico"
BACKUPS_DIR     = APP_DIR / "backups"
EXPORTS_DIR     = APP_DIR / "exports"
LOGS_DIR        = DATA_DIR / "logs"
SETTINGS_FILE   = DATA_DIR / "settings.json"
TODAY_SCANS_FILE= DATA_DIR / "today_scans.pkl"
ENCODINGS_FILE  = DATA_DIR / "encodings.pkl"
AUDIT_LOG       = LOGS_DIR / "attendance_log.csv"

# 3.0: scratch-space folders for the new Whiteboard and Notepad tabs.
WHITEBOARD_DIR  = APP_DIR / "Whiteboard"
NOTES_DIR       = APP_DIR / "Notes"

FIXED_EXCEL_FILENAME = "FaceAttend_Students.xlsx"
EXCEL_FILE = APP_DIR / FIXED_EXCEL_FILENAME

SOUND_ERROR        = APP_DIR / "sound_error.wav"
SOUND_CAMERA_START = APP_DIR / "sound_camera_start.wav"
SOUND_SUCCESS      = APP_DIR / "sound_success.mp3"    # positive actions: add student, etc.
SOUND_NEGATIVE      = APP_DIR / "sound_negative.mp3"   # negative/destructive actions: remove student
SOUND_STARTUP       = APP_DIR / "sound_startup.mp3"    # plays once when the app launches
SOUND_EXPORT        = APP_DIR / "sound_export.mp3"     # plays on successful Export AND Import

OLLAMA_URL   = os.getenv("FACEATTEND_OLLAMA_URL",   "http://localhost:11434")
OLLAMA_MODEL = os.getenv("FACEATTEND_OLLAMA_MODEL", "qwen2.5:0.5b")

# ── schema ───────────────────────────────────────────────────────────────────
COLUMNS = [
    "Sl.No", "Roll No", "UID", "Name", "Grade", "House",
    "Parent Phone", "Parent Email", "No. of Days Present",
    "XP", "Multiplier", "Streak",
]
PHOTO_COLUMN = len(COLUMNS) + 1
EXCEL_COLUMN_WIDTHS = {
    "Sl.No": 7, "Roll No": 13, "UID": 10, "Name": 26,
    "Grade": 10, "House": 12,
    "Parent Phone": 18, "Parent Email": 30,
    "No. of Days Present": 20, "XP": 10, "Multiplier": 12, "Streak": 10,
}
HOUSE_COLORS = {
    "Red":    "#ef4444",
    "Blue":   "#3b82f6",
    "Green":  "#22c55e",
    "Orange": "#f97316",
    "None":   "",
}
HOUSES = ["None", "Red", "Blue", "Green", "Orange"]

LATE_TIME      = dtime(8, 32)
FACE_THRESHOLD = 0.50
FRAME_SKIP     = 3
CAMERA_W       = 640
CAMERA_H       = 480

# ── global Excel lock ────────────────────────────────────────────────────────
_excel_lock = threading.Lock()

# ═══════════════════════════════════════════════════════════════════════════════
#  COLOUR PALETTE  (dark)
# ═══════════════════════════════════════════════════════════════════════════════
def play_sound(path: Path) -> None:
    """Play an MP3 file via pygame.mixer — non-blocking, daemon thread."""
    if not path.exists():
        return

    def _play():
        try:
            import pygame
            pygame.mixer.init()
            pygame.mixer.music.load(str(path))
            pygame.mixer.music.play()
            import time as _time
            while pygame.mixer.music.get_busy():
                _time.sleep(0.05)
        except Exception as e:
            print(f"[FaceAttend] Sound error: {e}")

    threading.Thread(target=_play, daemon=True).start()
class _PaletteDark:
    rail        = "#0a0e15"
    bg          = "#0e1117"
    sidebar     = "#12161f"
    sidebar2    = "#1a2130"
    surface     = "#1c2233"
    surface2    = "#222b3a"
    hover       = "#283248"
    border      = "#2a3448"
    border_soft = "#202838"
    text        = "#e8edf5"
    muted       = "#7b8aa0"
    faint       = "#4d586c"
    blue        = "#3b82f6"
    blue_dark   = "#2563eb"
    blue_glow   = "#1d4ed8"
    blue_soft   = "#132236"
    green       = "#22c55e"
    green_dark  = "#16a34a"
    green_soft  = "#0d2218"
    amber       = "#f59e0b"
    amber_soft  = "#241b08"
    red         = "#ef4444"
    red_dark    = "#dc2626"
    red_soft    = "#2a1316"
    purple      = "#a855f7"
    purple_soft = "#1c1430"
    ink         = "#f1f5f9"
    tint_blue   = "#0f1e35"
    tint_green  = "#0d2218"
    tint_amber  = "#231a05"
    tint_purple = "#1a1030"
    chart_bg    = "#0e1117"
    chart_fg    = "#e8edf5"
    chart_grid  = "#2a3448"
    calendar_present = "#22c55e"
    calendar_absent  = "#ef4444"
    calendar_neutral = "#2a3448"
    font_family = "Segoe UI"
    style_name  = "Nordic Slate"
    glyph       = "\u2B21"


class _PaletteLight:
    rail        = "#f1f5f9"
    bg          = "#ffffff"
    sidebar     = "#f8fafc"
    sidebar2    = "#f1f5f9"
    surface     = "#ffffff"
    surface2    = "#f8fafc"
    hover       = "#e2e8f0"
    border      = "#cbd5e1"
    border_soft = "#e2e8f0"
    text        = "#0f172a"
    muted       = "#64748b"
    faint       = "#94a3b8"
    blue        = "#2563eb"
    blue_dark   = "#1d4ed8"
    blue_glow   = "#1e40af"
    blue_soft   = "#dbeafe"
    green       = "#16a34a"
    green_dark  = "#15803d"
    green_soft  = "#dcfce7"
    amber       = "#d97706"
    amber_soft  = "#fef3c7"
    red         = "#dc2626"
    red_dark    = "#b91c1c"
    red_soft    = "#fee2e2"
    purple      = "#7c3aed"
    purple_soft = "#ede9fe"
    ink         = "#0f172a"
    tint_blue   = "#eff6ff"
    tint_green  = "#f0fdf4"
    tint_amber  = "#fffbeb"
    tint_purple = "#f5f3ff"
    chart_bg    = "#ffffff"
    chart_fg    = "#0f172a"
    chart_grid  = "#e2e8f0"
    calendar_present = "#16a34a"
    calendar_absent  = "#dc2626"
    calendar_neutral = "#e2e8f0"
    font_family = "Segoe UI"
    style_name  = "Nordic Slate"
    glyph       = "\u2B21"


P = _PaletteDark()  # module-level active palette

# ═══════════════════════════════════════════════════════════════════════════════
#  13 LANGUAGE-KEYED VISUAL IDENTITIES
#  Each supported language carries its own accent-colour family, font
#  pairing, and rail glyph, so switching the language in Settings genuinely
#  changes how the app looks after the automatic restart — not just which
#  words appear on the labels. Names are abstract palette names (not tied
#  to any ethnicity or culture) — only the language *code* selects one.
# ═══════════════════════════════════════════════════════════════════════════════
THEME_VARIANTS = {
    "en": {"style_name": "Nordic Slate",    "accent": "#3b82f6", "accent_dark": "#2563eb", "accent_glow": "#1d4ed8", "soft_dark": "#132236", "soft_light": "#dbeafe", "font": "Segoe UI",   "glyph": "\u2B21"},
    "hi": {"style_name": "Ember Coral",     "accent": "#f97316", "accent_dark": "#ea580c", "accent_glow": "#c2410c", "soft_dark": "#2a1708", "soft_light": "#ffedd5", "font": "Nirmala UI", "glyph": "\u25C6"},
    "bn": {"style_name": "Meadow Teal",     "accent": "#14b8a6", "accent_dark": "#0d9488", "accent_glow": "#0f766e", "soft_dark": "#0b2320", "soft_light": "#ccfbf1", "font": "Nirmala UI", "glyph": "\u25C8"},
    "te": {"style_name": "Orchid Dusk",     "accent": "#a855f7", "accent_dark": "#9333ea", "accent_glow": "#7e22ce", "soft_dark": "#1e1330", "soft_light": "#f3e8ff", "font": "Nirmala UI", "glyph": "\u2B22"},
    "mr": {"style_name": "Harbor Blue",     "accent": "#0ea5e9", "accent_dark": "#0284c7", "accent_glow": "#0369a1", "soft_dark": "#0c2233", "soft_light": "#e0f2fe", "font": "Nirmala UI", "glyph": "\u2B23"},
    "ta": {"style_name": "Copper Ridge",    "accent": "#d97706", "accent_dark": "#b45309", "accent_glow": "#92400e", "soft_dark": "#241a08", "soft_light": "#fef3c7", "font": "Nirmala UI", "glyph": "\u2B1F"},
    "gu": {"style_name": "Violet Storm",    "accent": "#8b5cf6", "accent_dark": "#7c3aed", "accent_glow": "#6d28d9", "soft_dark": "#1c1730", "soft_light": "#ede9fe", "font": "Nirmala UI", "glyph": "\u2B21"},
    "ur": {"style_name": "Golden Hour",     "accent": "#eab308", "accent_dark": "#ca8a04", "accent_glow": "#a16207", "soft_dark": "#251f06", "soft_light": "#fef9c3", "font": "Nirmala UI", "glyph": "\u25C7"},
    "kn": {"style_name": "Arctic Frost",    "accent": "#38bdf8", "accent_dark": "#0ea5e9", "accent_glow": "#0284c7", "soft_dark": "#0b1f2e", "soft_light": "#e0f2fe", "font": "Nirmala UI", "glyph": "\u2B20"},
    "or": {"style_name": "Rosewood",        "accent": "#f43f5e", "accent_dark": "#e11d48", "accent_glow": "#be123c", "soft_dark": "#2a1116", "soft_light": "#ffe4e6", "font": "Nirmala UI", "glyph": "\u25C6"},
    "ml": {"style_name": "Emerald Deep",    "accent": "#10b981", "accent_dark": "#059669", "accent_glow": "#047857", "soft_dark": "#08211a", "soft_light": "#d1fae5", "font": "Nirmala UI", "glyph": "\u2B22"},
    "pa": {"style_name": "Solar Amber",     "accent": "#f59e0b", "accent_dark": "#d97706", "accent_glow": "#b45309", "soft_dark": "#241b08", "soft_light": "#fef3c7", "font": "Nirmala UI", "glyph": "\u25C8"},
    "as": {"style_name": "Midnight Indigo", "accent": "#6366f1", "accent_dark": "#4f46e5", "accent_glow": "#4338ca", "soft_dark": "#181b30", "soft_light": "#e0e7ff", "font": "Nirmala UI", "glyph": "\u2B23"},
}


def set_theme(dark: bool = True, lang: str = "en"):
    """
    Swap the global active palette between Dark and Light, AND apply the
    given language's visual identity on top of it (accent colour family,
    font pairing, rail glyph). This is what makes changing the language
    in Settings feel like switching to a differently-designed app once
    FaceAttend restarts, instead of only swapping label text.
    """
    global P
    base = _PaletteDark() if dark else _PaletteLight()
    variant = THEME_VARIANTS.get(lang, THEME_VARIANTS["en"])
    accent_soft = variant["soft_dark"] if dark else variant["soft_light"]
    base.blue       = variant["accent"]
    base.blue_dark  = variant["accent_dark"]
    base.blue_glow  = variant["accent_glow"]
    base.blue_soft  = accent_soft
    base.tint_blue   = accent_soft
    base.font_family = variant["font"]
    base.style_name  = variant["style_name"]
    base.glyph       = variant["glyph"]
    P = base
    # Re-bind the shared font tuples to this language's font pairing too,
    # so headings/body/small text across every tab pick it up as soon as
    # the UI is (re)built — this runs before _build() on every launch.
    global F_H1, F_H2, F_H3, F_BODY, F_BODY_B, F_SMALL, F_SMALL_B
    fam = variant["font"]
    F_H1      = (fam, 22, "bold")
    F_H2      = (fam, 15, "bold")
    F_H3      = (fam, 12, "bold")
    F_BODY    = (fam, 10)
    F_BODY_B  = (fam, 10, "bold")
    F_SMALL   = (fam, 9)
    F_SMALL_B = (fam, 9, "bold")
    return P


# ═══════════════════════════════════════════════════════════════════════════════
#  INTERNATIONALISATION (i18n)
# ═══════════════════════════════════════════════════════════════════════════════
TRANSLATIONS = {
    "en": {
        "app_title": "FaceAttend {ver}  ·  Smart Attendance System",
        "home_scanner": "Home & Scanner",
        "analytics": "Analytics",
        "ai_buddy": "AI Buddy",
        "messages": "Messages",
        "houses": "House Competition",
        "console": "Console",
        "settings": "Settings",
        "start_scan": "▶  Start Scan",
        "stop_scan": "■  Stop Scan",
        "refresh": "↺  Refresh",
        "add_student": "➕  Add",
        "edit_student": "✏  Edit",
        "remove_student": "🗑  Remove",
        "import": "📥  Import",
        "export": "⬇  Export",
        "total_students": "Total Students",
        "present_today": "Present Today",
        "absent_today": "Absent Today",
        "xp_leader": "XP Leader",
        "enrolled": "Enrolled",
        "scanned_in": "Scanned in",
        "not_yet_seen": "Not yet seen",
        "top_scorer": "Top scorer",
        "teacher_dashboard": "Teacher Dashboard",
        "search": "Search students...",
        "clear": "Clear",
        "last_recognition": "LAST RECOGNITION",
        "waiting": "Waiting for student…",
        "start_scanning": "Start scanning to mark attendance.",
        "activity_feed": "Activity Feed",
        "no_activity": "No activity yet — add a student or start scanning.",
        "camera_off": "📷  Camera is off",
        "settings_saved": "Settings saved.",
        "settings_reset": "Settings reset to defaults.",
        "backup_saved": "Backup saved.",
        "email_sent": "Email sent to parent of {name}",
        "whatsapp_opened": "WhatsApp opened for parent of {name}",
        "student_added": "{name} added successfully!",
        "student_updated": "{name} updated.",
        "student_removed": "{name} removed.",
        "select_student_first": "Select a student row first.",
        "attendance_late": "Late",
        "attendance_ontime": "On time",
        "welcome": "Welcome, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "Language",
        "theme": "Theme",
        "dark": "Dark",
        "light": "Light",
        "schedule": "Schedule",
        "badges": "Badges & Achievements",
        "badge_first_steps": "First Steps",
        "badge_week_warrior": "Week Warrior",
        "badge_month_master": "Month Master",
        "badge_early_bird": "Early Bird",
        "badge_centurion": "Centurion",
        "badge_house_hero": "House Hero",
        "qr_code": "QR Code",
        "generate_qr": "Generate QR",
        "scan_qr": "Scan QR",
        "liveness": "Liveness Detection",
        "voice_feedback": "Voice Feedback",
        "multi_camera": "Multi-Camera",
        "camera_select": "Select Camera",
        "api_server": "REST API Server",
        "api_running": "API running on port {port}",
        "api_stopped": "API stopped",
        "calendar": "Attendance Calendar",
        "jan": "Jan", "feb": "Feb", "mar": "Mar", "apr": "Apr",
        "may": "May", "jun": "Jun", "jul": "Jul", "aug": "Aug",
        "sep": "Sep", "oct": "Oct", "nov": "Nov", "dec": "Dec",
        "mon": "Mon", "tue": "Tue", "wed": "Wed", "thu": "Thu",
        "fri": "Fri", "sat": "Sat", "sun": "Sun",
        "whiteboard": "Whiteboard", "notepad": "Notepad",
        "mark_attendance": "Mark Attendance", "change_section": "Change Section", "select_all": "Select All",
    },
    "hi": {
        "app_title": "FaceAttend {ver}  ·  स्मार्ट उपस्थिति प्रणाली",
        "home_scanner": "होम और स्कैनर",
        "analytics": "विश्लेषण",
        "ai_buddy": "AI सहायक",
        "messages": "संदेश",
        "houses": "हाउस प्रतियोगिता",
        "console": "कंसोल",
        "settings": "सेटिंग्स",
        "start_scan": "▶  स्कैन शुरू करें",
        "stop_scan": "■  स्कैन रोकें",
        "refresh": "↺  रिफ्रेश",
        "add_student": "➕  जोड़ें",
        "edit_student": "✏  संपादित करें",
        "remove_student": "🗑  हटाएं",
        "import": "📥  आयात",
        "export": "⬇  निर्यात",
        "total_students": "कुल छात्र",
        "present_today": "आज उपस्थित",
        "absent_today": "आज अनुपस्थित",
        "xp_leader": "XP लीडर",
        "enrolled": "नामांकित",
        "scanned_in": "स्कैन किए",
        "not_yet_seen": "अभी तक नहीं",
        "top_scorer": "शीर्ष स्कोरर",
        "teacher_dashboard": "शिक्षक डैशबोर्ड",
        "search": "छात्र खोजें...",
        "clear": "साफ करें",
        "last_recognition": "अंतिम पहचान",
        "waiting": "छात्र का इंतजार…",
        "start_scanning": "उपस्थिति दर्ज करने के लिए स्कैनिंग शुरू करें।",
        "activity_feed": "गतिविधि फ़ीड",
        "no_activity": "अभी तक कोई गतिविधि नहीं — छात्र जोड़ें या स्कैनिंग शुरू करें।",
        "camera_off": "📷  कैमरा बंद है",
        "settings_saved": "सेटिंग्स सहेजी गईं।",
        "settings_reset": "सेटिंग्स डिफ़ॉल्ट पर रीसेट।",
        "backup_saved": "बैकअप सहेजा गया।",
        "email_sent": "{name} के माता-पिता को ईमेल भेजा गया",
        "whatsapp_opened": "{name} के माता-पिता के लिए WhatsApp खोला गया",
        "student_added": "{name} सफलतापूर्वक जोड़ा गया!",
        "student_updated": "{name} अपडेट किया गया।",
        "student_removed": "{name} हटाया गया।",
        "select_student_first": "पहले एक छात्र चुनें।",
        "attendance_late": "देर",
        "attendance_ontime": "समय पर",
        "welcome": "स्वागत है, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "भाषा",
        "theme": "थीम",
        "dark": "डार्क",
        "light": "लाइट",
        "schedule": "अनुसूची",
        "badges": "बैज और उपलब्धियाँ",
        "badge_first_steps": "पहले कदम",
        "badge_week_warrior": "सप्ताह योद्धा",
        "badge_month_master": "माह मास्टर",
        "badge_early_bird": "सुबह का पंछी",
        "badge_centurion": "सैनिक",
        "badge_house_hero": "हाउस हीरो",
        "qr_code": "QR कोड",
        "generate_qr": "QR बनाएं",
        "scan_qr": "QR स्कैन करें",
        "liveness": "जीवितता पहचान",
        "voice_feedback": "आवाज़ प्रतिक्रिया",
        "multi_camera": "मल्टी-कैमरा",
        "camera_select": "कैमरा चुनें",
        "api_server": "REST API सर्वर",
        "api_running": "पोर्ट {port} पर API चल रहा है",
        "api_stopped": "API बंद",
        "calendar": "उपस्थिति कैलेंडर",
        "jan": "जन", "feb": "फर", "mar": "मार्च", "apr": "अप्रै",
        "may": "मई", "jun": "जून", "jul": "जुल", "aug": "अग",
        "sep": "सित", "oct": "अक्टू", "nov": "नव", "dec": "दिस",
        "mon": "सोम", "tue": "मंगल", "wed": "बुध", "thu": "गुरु",
        "mark_attendance": "उपस्थिति दर्ज करें", "change_section": "सेक्शन बदलें", "select_all": "सभी चुनें",
        "whiteboard": "व्हाइटबोर्ड", "notepad": "नोटपैड",
        "fri": "शुक्र", "sat": "शनि", "sun": "रवि",
    },
    "bn": {
        "app_title": "FaceAttend {ver}  ·  স্মার্ট উপস্থিতি ব্যবস্থা",
        "home_scanner": "হোম ও স্ক্যানার",
        "analytics": "বিশ্লেষণ",
        "ai_buddy": "AI সহায়ক",
        "messages": "বার্তা",
        "houses": "হাউস প্রতিযোগিতা",
        "console": "কনসোল",
        "settings": "সেটিংস",
        "start_scan": "▶  স্ক্যান শুরু করুন",
        "stop_scan": "■  স্ক্যান বন্ধ করুন",
        "refresh": "↺  রিফ্রেশ",
        "add_student": "➕  যুক্ত করুন",
        "edit_student": "✏  সম্পাদনা",
        "remove_student": "🗑  অপসারণ",
        "import": "📥  আমদানি",
        "export": "⬇  রপ্তানি",
        "total_students": "মোট শিক্ষার্থী",
        "present_today": "আজ উপস্থিত",
        "absent_today": "আজ অনুপস্থিত",
        "xp_leader": "XP লিডার",
        "enrolled": "ভর্তি হয়েছে",
        "scanned_in": "স্ক্যান হয়েছে",
        "not_yet_seen": "এখনো দেখা যায়নি",
        "top_scorer": "সর্বোচ্চ স্কোরার",
        "teacher_dashboard": "শিক্ষক ড্যাশবোর্ড",
        "search": "শিক্ষার্থী খুঁজুন...",
        "clear": "পরিষ্কার",
        "last_recognition": "সর্বশেষ স্বীকৃতি",
        "waiting": "শিক্ষার্থীর জন্য অপেক্ষা…",
        "start_scanning": "উপস্থিতি নথিভুক্ত করতে স্ক্যান শুরু করুন।",
        "activity_feed": "কার্যকলাপ ফিড",
        "no_activity": "এখনও কোনো কার্যকলাপ নেই — শিক্ষার্থী যুক্ত করুন বা স্ক্যান শুরু করুন।",
        "camera_off": "📷  ক্যামেরা বন্ধ আছে",
        "settings_saved": "সেটিংস সংরক্ষিত হয়েছে।",
        "settings_reset": "সেটিংস ডিফল্টে রিসেট হয়েছে।",
        "backup_saved": "ব্যাকআপ সংরক্ষিত হয়েছে।",
        "email_sent": "{name} এর অভিভাবককে ইমেইল পাঠানো হয়েছে",
        "whatsapp_opened": "{name} এর অভিভাবকের জন্য WhatsApp খোলা হয়েছে",
        "student_added": "{name} সফলভাবে যুক্ত হয়েছে!",
        "student_updated": "{name} আপডেট হয়েছে।",
        "student_removed": "{name} অপসারিত হয়েছে।",
        "select_student_first": "প্রথমে একজন শিক্ষার্থী নির্বাচন করুন।",
        "attendance_late": "বিলম্বিত",
        "attendance_ontime": "সময়মতো",
        "welcome": "স্বাগতম, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "ভাষা",
        "theme": "থিম",
        "dark": "ডার্ক",
        "light": "লাইট",
        "schedule": "সময়সূচী",
        "badges": "ব্যাজ ও কৃতিত্ব",
        "badge_first_steps": "প্রথম পদক্ষেপ",
        "badge_week_warrior": "সপ্তাহের যোদ্ধা",
        "badge_month_master": "মাসের মাস্টার",
        "badge_early_bird": "প্রভাতী পাখি",
        "badge_centurion": "শতপতি",
        "badge_house_hero": "হাউস হিরো",
        "qr_code": "QR কোড",
        "generate_qr": "QR তৈরি করুন",
        "scan_qr": "QR স্ক্যান করুন",
        "liveness": "লাইভনেস সনাক্তকরণ",
        "voice_feedback": "ভয়েস প্রতিক্রিয়া",
        "multi_camera": "মাল্টি-ক্যামেরা",
        "camera_select": "ক্যামেরা নির্বাচন করুন",
        "api_server": "REST API সার্ভার",
        "api_running": "পোর্ট {port} এ API চলছে",
        "api_stopped": "API বন্ধ হয়েছে",
        "calendar": "উপস্থিতি ক্যালেন্ডার",
        "jan": "জানু",
        "feb": "ফেব্রু",
        "mar": "মার্চ",
        "apr": "এপ্রিল",
        "may": "মে",
        "jun": "জুন",
        "jul": "জুলাই",
        "aug": "আগস্ট",
        "sep": "সেপ্ট",
        "oct": "অক্টো",
        "nov": "নভে",
        "dec": "ডিসে",
        "mon": "সোম",
        "tue": "মঙ্গল",
        "wed": "বুধ",
        "thu": "বৃহঃ",
        "mark_attendance": "উপস্থিতি নথিভুক্ত করুন", "change_section": "সেকশন পরিবর্তন করুন", "select_all": "সব নির্বাচন করুন",
        "fri": "শুক্র",
        "whiteboard": "হোয়াইটবোর্ড", "notepad": "নোটপ্যাড",
        "sat": "শনি",
        "sun": "রবি",
    },
    "te": {
        "app_title": "FaceAttend {ver}  ·  స్మార్ట్ హాజరు వ్యవస్థ",
        "home_scanner": "హోమ్ & స్కానర్",
        "analytics": "విశ్లేషణలు",
        "ai_buddy": "AI బడ్డీ",
        "messages": "సందేశాలు",
        "houses": "హౌస్ పోటీ",
        "console": "కన్సోల్",
        "settings": "సెట్టింగ్‌లు",
        "start_scan": "▶  స్కాన్ ప్రారంభించండి",
        "stop_scan": "■  స్కాన్ ఆపండి",
        "refresh": "↺  రిఫ్రెష్",
        "add_student": "➕  జోడించండి",
        "edit_student": "✏  సవరించండి",
        "remove_student": "🗑  తీసివేయండి",
        "import": "📥  దిగుమతి",
        "export": "⬇  ఎగుమతి",
        "total_students": "మొత్తం విద్యార్థులు",
        "present_today": "ఈ రోజు హాజరు",
        "absent_today": "ఈ రోజు గైర్హాజరు",
        "xp_leader": "XP లీడర్",
        "enrolled": "నమోదు చేయబడింది",
        "scanned_in": "స్కాన్ చేయబడింది",
        "not_yet_seen": "ఇంకా చూడలేదు",
        "top_scorer": "టాప్ స్కోరర్",
        "teacher_dashboard": "టీచర్ డాష్‌బోర్డ్",
        "search": "విద్యార్థులను శోధించండి...",
        "clear": "క్లియర్",
        "last_recognition": "చివరి గుర్తింపు",
        "waiting": "విద్యార్థి కోసం వేచి ఉంది…",
        "start_scanning": "హాజరు నమోదు చేయడానికి స్కానింగ్ ప్రారంభించండి।",
        "activity_feed": "కార్యకలాప ఫీడ్",
        "no_activity": "ఇంకా కార్యకలాపం లేదు — విద్యార్థిని జోడించండి లేదా స్కానింగ్ ప్రారంభించండి।",
        "camera_off": "📷  కెమెరా ఆఫ్‌లో ఉంది",
        "settings_saved": "సెట్టింగ్‌లు సేవ్ చేయబడ్డాయి।",
        "settings_reset": "సెట్టింగ్‌లు డిఫాల్ట్‌కు రీసెట్ చేయబడ్డాయి।",
        "backup_saved": "బ్యాకప్ సేవ్ చేయబడింది।",
        "email_sent": "{name} యొక్క తల్లిదండ్రులకు ఇమెయిల్ పంపబడింది",
        "whatsapp_opened": "{name} యొక్క తల్లిదండ్రుల కోసం WhatsApp తెరవబడింది",
        "student_added": "{name} విజయవంతంగా జోడించబడింది!",
        "student_updated": "{name} నవీకరించబడింది।",
        "student_removed": "{name} తీసివేయబడింది।",
        "select_student_first": "మొదట ఒక విద్యార్థిని ఎంచుకోండి।",
        "attendance_late": "ఆలస్యం",
        "attendance_ontime": "సమయానికి",
        "welcome": "స్వాగతం, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "భాష",
        "theme": "థీమ్",
        "dark": "డార్క్",
        "light": "లైట్",
        "schedule": "షెడ్యూల్",
        "badges": "బ్యాడ్జీలు & సాధనలు",
        "badge_first_steps": "మొదటి అడుగులు",
        "badge_week_warrior": "వీక్ వారియర్",
        "badge_month_master": "మంత్ మాస్టర్",
        "badge_early_bird": "ఎర్లీ బర్డ్",
        "badge_centurion": "సెంచూరియన్",
        "badge_house_hero": "హౌస్ హీరో",
        "qr_code": "QR కోడ్",
        "generate_qr": "QR తయారు చేయండి",
        "scan_qr": "QR స్కాన్ చేయండి",
        "liveness": "లైవ్‌నెస్ గుర్తింపు",
        "voice_feedback": "వాయిస్ ఫీడ్‌బ్యాక్",
        "multi_camera": "మల్టీ-కెమెరా",
        "camera_select": "కెమెరాను ఎంచుకోండి",
        "api_server": "REST API సర్వర్",
        "api_running": "పోర్ట్ {port} లో API నడుస్తోంది",
        "api_stopped": "API ఆగిపోయింది",
        "calendar": "హాజరు క్యాలెండర్",
        "jan": "జన",
        "feb": "ఫెబ్ర",
        "mar": "మార్చి",
        "apr": "ఏప్రి",
        "may": "మే",
        "jun": "జూన్",
        "jul": "జూలై",
        "aug": "ఆగ",
        "sep": "సెప్టెం",
        "oct": "అక్టో",
        "nov": "నవం",
        "dec": "డిసెం",
        "mon": "సోమ",
        "tue": "మంగళ",
        "mark_attendance": "హాజరు నమోదు చేయండి", "change_section": "సెక్షన్ మార్చండి", "select_all": "అన్నీ ఎంచుకోండి",
        "wed": "బుధ",
        "thu": "గురు",
        "whiteboard": "వైట్‌బోర్డ్", "notepad": "నోట్‌ప్యాడ్",
        "fri": "శుక్ర",
        "sat": "శని",
        "sun": "ఆది",
    },
    "mr": {
        "app_title": "FaceAttend {ver}  ·  स्मार्ट उपस्थिती प्रणाली",
        "home_scanner": "होम आणि स्कॅनर",
        "analytics": "विश्लेषण",
        "ai_buddy": "AI सहाय्यक",
        "messages": "संदेश",
        "houses": "हाऊस स्पर्धा",
        "console": "कन्सोल",
        "settings": "सेटिंग्ज",
        "start_scan": "▶  स्कॅन सुरू करा",
        "stop_scan": "■  स्कॅन थांबवा",
        "refresh": "↺  रिफ्रेश",
        "add_student": "➕  जोडा",
        "edit_student": "✏  संपादित करा",
        "remove_student": "🗑  काढून टाका",
        "import": "📥  आयात",
        "export": "⬇  निर्यात",
        "total_students": "एकूण विद्यार्थी",
        "present_today": "आज उपस्थित",
        "absent_today": "आज अनुपस्थित",
        "xp_leader": "XP लीडर",
        "enrolled": "नोंदणीकृत",
        "scanned_in": "स्कॅन केले",
        "not_yet_seen": "अद्याप दिसले नाही",
        "top_scorer": "सर्वोच्च गुणधारक",
        "teacher_dashboard": "शिक्षक डॅशबोर्ड",
        "search": "विद्यार्थी शोधा...",
        "clear": "साफ करा",
        "last_recognition": "शेवटची ओळख",
        "waiting": "विद्यार्थ्याची वाट पाहत आहे…",
        "start_scanning": "उपस्थिती नोंदवण्यासाठी स्कॅनिंग सुरू करा.",
        "activity_feed": "क्रियाकलाप फीड",
        "no_activity": "अद्याप कोणतीही क्रिया नाही — विद्यार्थी जोडा किंवा स्कॅनिंग सुरू करा.",
        "camera_off": "📷  कॅमेरा बंद आहे",
        "settings_saved": "सेटिंग्ज जतन केल्या.",
        "settings_reset": "सेटिंग्ज डीफॉल्टवर रीसेट केल्या.",
        "backup_saved": "बॅकअप जतन केला.",
        "email_sent": "{name} च्या पालकांना ईमेल पाठवला",
        "whatsapp_opened": "{name} च्या पालकांसाठी WhatsApp उघडले",
        "student_added": "{name} यशस्वीरित्या जोडले!",
        "student_updated": "{name} अपडेट केले.",
        "student_removed": "{name} काढून टाकले.",
        "select_student_first": "प्रथम एक विद्यार्थी निवडा.",
        "attendance_late": "उशीर",
        "attendance_ontime": "वेळेवर",
        "welcome": "स्वागत आहे, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "भाषा",
        "theme": "थीम",
        "dark": "डार्क",
        "light": "लाइट",
        "schedule": "वेळापत्रक",
        "badges": "बॅज आणि कामगिरी",
        "badge_first_steps": "पहिले पाऊल",
        "badge_week_warrior": "आठवड्याचा योद्धा",
        "badge_month_master": "महिन्याचा मास्टर",
        "badge_early_bird": "अर्ली बर्ड",
        "badge_centurion": "सेंच्युरियन",
        "badge_house_hero": "हाऊस हिरो",
        "qr_code": "QR कोड",
        "generate_qr": "QR तयार करा",
        "scan_qr": "QR स्कॅन करा",
        "liveness": "लाइव्हनेस तपासणी",
        "voice_feedback": "व्हॉइस फीडबॅक",
        "multi_camera": "मल्टी-कॅमेरा",
        "camera_select": "कॅमेरा निवडा",
        "api_server": "REST API सर्व्हर",
        "api_running": "पोर्ट {port} वर API चालू आहे",
        "api_stopped": "API थांबले",
        "calendar": "उपस्थिती कॅलेंडर",
        "jan": "जाने",
        "feb": "फेब्रु",
        "mar": "मार्च",
        "apr": "एप्रि",
        "may": "मे",
        "jun": "जून",
        "jul": "जुलै",
        "aug": "ऑगस्ट",
        "sep": "सप्टें",
        "oct": "ऑक्टो",
        "nov": "नोव्हें",
        "dec": "डिसें",
        "mark_attendance": "उपस्थिती नोंदवा", "change_section": "विभाग बदला", "select_all": "सर्व निवडा",
        "mon": "सोम",
        "tue": "मंगळ",
        "wed": "बुध",
        "whiteboard": "व्हाइटबोर्ड", "notepad": "नोटपॅड",
        "thu": "गुरु",
        "fri": "शुक्र",
        "sat": "शनि",
        "sun": "रवि",
    },
    "ta": {
        "app_title": "FaceAttend {ver}  ·  ஸ்மார்ட் வருகை அமைப்பு",
        "home_scanner": "முகப்பு & ஸ்கேனர்",
        "analytics": "பகுப்பாய்வு",
        "ai_buddy": "AI துணை",
        "messages": "செய்திகள்",
        "houses": "ஹவுஸ் போட்டி",
        "console": "கன்சோல்",
        "settings": "அமைப்புகள்",
        "start_scan": "▶  ஸ்கேன் தொடங்கு",
        "stop_scan": "■  ஸ்கேன் நிறுத்து",
        "refresh": "↺  புதுப்பிக்க",
        "add_student": "➕  சேர்க்க",
        "edit_student": "✏  திருத்த",
        "remove_student": "🗑  அகற்ற",
        "import": "📥  இறக்குமதி",
        "export": "⬇  ஏற்றுமதி",
        "total_students": "மொத்த மாணவர்கள்",
        "present_today": "இன்று வருகை",
        "absent_today": "இன்று வராதவர்",
        "xp_leader": "XP தலைவர்",
        "enrolled": "பதிவு செய்யப்பட்டது",
        "scanned_in": "ஸ்கேன் செய்யப்பட்டது",
        "not_yet_seen": "இன்னும் காணப்படவில்லை",
        "top_scorer": "சிறந்த மதிப்பெண் பெற்றவர்",
        "teacher_dashboard": "ஆசிரியர் டாஷ்போர்டு",
        "search": "மாணவர்களைத் தேடு...",
        "clear": "அழி",
        "last_recognition": "கடைசி அங்கீகாரம்",
        "waiting": "மாணவருக்காக காத்திருக்கிறது…",
        "start_scanning": "வருகையைப் பதிவு செய்ய ஸ்கேனிங் தொடங்கவும்.",
        "activity_feed": "செயல்பாட்டு ஊட்டம்",
        "no_activity": "இன்னும் செயல்பாடு இல்லை — மாணவரைச் சேர்க்கவும் அல்லது ஸ்கேன் தொடங்கவும்.",
        "camera_off": "📷  கேமரா ஆஃப் செய்யப்பட்டது",
        "settings_saved": "அமைப்புகள் சேமிக்கப்பட்டன.",
        "settings_reset": "அமைப்புகள் இயல்புநிலைக்கு மீட்டமைக்கப்பட்டன.",
        "backup_saved": "காப்புப்பிரதி சேமிக்கப்பட்டது.",
        "email_sent": "{name} இன் பெற்றோருக்கு மின்னஞ்சல் அனுப்பப்பட்டது",
        "whatsapp_opened": "{name} இன் பெற்றோருக்காக WhatsApp திறக்கப்பட்டது",
        "student_added": "{name} வெற்றிகரமாக சேர்க்கப்பட்டார்!",
        "student_updated": "{name} புதுப்பிக்கப்பட்டது.",
        "student_removed": "{name} அகற்றப்பட்டது.",
        "select_student_first": "முதலில் ஒரு மாணவரைத் தேர்ந்தெடுக்கவும்.",
        "attendance_late": "தாமதம்",
        "attendance_ontime": "சரியான நேரம்",
        "welcome": "வரவேற்கிறோம், {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "மொழி",
        "theme": "தீம்",
        "dark": "டார்க்",
        "light": "லைட்",
        "schedule": "அட்டவணை",
        "badges": "பேட்ஜ்கள் & சாதனைகள்",
        "badge_first_steps": "முதல் படிகள்",
        "badge_week_warrior": "வார வாரியர்",
        "badge_month_master": "மாத மாஸ்டர்",
        "badge_early_bird": "எர்லி பேர்ட்",
        "badge_centurion": "சென்சூரியன்",
        "badge_house_hero": "ஹவுஸ் ஹீரோ",
        "qr_code": "QR குறியீடு",
        "generate_qr": "QR உருவாக்கு",
        "scan_qr": "QR ஸ்கேன் செய்",
        "liveness": "லைவ்நெஸ் கண்டறிதல்",
        "voice_feedback": "குரல் பின்னூட்டம்",
        "multi_camera": "பல கேமரா",
        "camera_select": "கேமராவைத் தேர்ந்தெடு",
        "api_server": "REST API சேவையகம்",
        "api_running": "போர்ட் {port} இல் API இயங்குகிறது",
        "api_stopped": "API நிறுத்தப்பட்டது",
        "calendar": "வருகை காலண்டர்",
        "jan": "ஜன",
        "feb": "பிப்ர",
        "mar": "மார்ச்",
        "apr": "ஏப்ரல்",
        "may": "மே",
        "jun": "ஜூன்",
        "jul": "ஜூலை",
        "aug": "ஆக",
        "sep": "செப்",
        "oct": "அக்",
        "mark_attendance": "வருகையை பதிவு செய்யவும்", "change_section": "பிரிவை மாற்று", "select_all": "அனைத்தையும் தேர்ந்தெடு",
        "nov": "நவ",
        "dec": "டிச",
        "mon": "திங்",
        "tue": "செவ்",
        "whiteboard": "வொயிட்போர்டு", "notepad": "குறிப்பேடு",
        "wed": "புத",
        "thu": "வியா",
        "fri": "வெள்",
        "sat": "சனி",
        "sun": "ஞாயி",
    },
    "gu": {
        "app_title": "FaceAttend {ver}  ·  સ્માર્ટ હાજરી સિસ્ટમ",
        "home_scanner": "હોમ અને સ્કેનર",
        "analytics": "વિશ્લેષણ",
        "ai_buddy": "AI સહાયક",
        "messages": "સંદેશા",
        "houses": "હાઉસ સ્પર્ધા",
        "console": "કન્સોલ",
        "settings": "સેટિંગ્સ",
        "start_scan": "▶  સ્કેન શરૂ કરો",
        "stop_scan": "■  સ્કેન બંધ કરો",
        "refresh": "↺  રિફ્રેશ",
        "add_student": "➕  ઉમેરો",
        "edit_student": "✏  સંપાદિત કરો",
        "remove_student": "🗑  દૂર કરો",
        "import": "📥  આયાત",
        "export": "⬇  નિકાસ",
        "total_students": "કુલ વિદ્યાર્થીઓ",
        "present_today": "આજે હાજર",
        "absent_today": "આજે ગેરહાજર",
        "xp_leader": "XP લીડર",
        "enrolled": "નોંધાયેલ",
        "scanned_in": "સ્કેન કરેલ",
        "not_yet_seen": "હજુ સુધી જોયું નથી",
        "top_scorer": "ટોચનો સ્કોરર",
        "teacher_dashboard": "શિક્ષક ડેશબોર્ડ",
        "search": "વિદ્યાર્થીઓ શોધો...",
        "clear": "સાફ કરો",
        "last_recognition": "છેલ્લી ઓળખ",
        "waiting": "વિદ્યાર્થીની રાહ જોવાઈ રહી છે…",
        "start_scanning": "હાજરી નોંધવા માટે સ્કેનિંગ શરૂ કરો.",
        "activity_feed": "પ્રવૃત્તિ ફીડ",
        "no_activity": "હજુ સુધી કોઈ પ્રવૃત્તિ નથી — વિદ્યાર્થી ઉમેરો અથવા સ્કેનિંગ શરૂ કરો.",
        "camera_off": "📷  કેમેરા બંધ છે",
        "settings_saved": "સેટિંગ્સ સાચવેલ.",
        "settings_reset": "સેટિંગ્સ ડિફોલ્ટ પર રીસેટ.",
        "backup_saved": "બેકઅપ સાચવ્યું.",
        "email_sent": "{name} ના વાલીને ઈમેલ મોકલ્યો",
        "whatsapp_opened": "{name} ના વાલી માટે WhatsApp ખોલ્યું",
        "student_added": "{name} સફળતાપૂર્વક ઉમેરાયું!",
        "student_updated": "{name} અપડેટ થયું.",
        "student_removed": "{name} દૂર કરાયું.",
        "select_student_first": "પ્રથમ વિદ્યાર્થી પસંદ કરો.",
        "attendance_late": "મોડું",
        "attendance_ontime": "સમયસર",
        "welcome": "સ્વાગત છે, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "ભાષા",
        "theme": "થીમ",
        "dark": "ડાર્ક",
        "light": "લાઇટ",
        "schedule": "સમયપત્રક",
        "badges": "બેજ અને સિદ્ધિઓ",
        "badge_first_steps": "પ્રથમ પગલાં",
        "badge_week_warrior": "વીક વોરિયર",
        "badge_month_master": "મંથ માસ્ટર",
        "badge_early_bird": "અર્લી બર્ડ",
        "badge_centurion": "સેન્ચ્યુરિયન",
        "badge_house_hero": "હાઉસ હીરો",
        "qr_code": "QR કોડ",
        "generate_qr": "QR બનાવો",
        "scan_qr": "QR સ્કેન કરો",
        "liveness": "લાઇવનેસ શોધ",
        "voice_feedback": "વોઇસ પ્રતિસાદ",
        "multi_camera": "મલ્ટી-કેમેરા",
        "camera_select": "કેમેરા પસંદ કરો",
        "api_server": "REST API સર્વર",
        "api_running": "પોર્ટ {port} પર API ચાલે છે",
        "api_stopped": "API બંધ થયું",
        "calendar": "હાજરી કેલેન્ડર",
        "jan": "જાન્યુ",
        "feb": "ફેબ્રુ",
        "mar": "માર્ચ",
        "apr": "એપ્રિ",
        "may": "મે",
        "jun": "જૂન",
        "jul": "જુલા",
        "aug": "ઓગ",
        "mark_attendance": "હાજરી નોંધો", "change_section": "વિભાગ બદલો", "select_all": "બધા પસંદ કરો",
        "sep": "સપ્ટે",
        "oct": "ઓક્ટો",
        "nov": "નવે",
        "dec": "ડિસે",
        "mon": "સોમ",
        "whiteboard": "વ્હાઇટબોર્ડ", "notepad": "નોટપેડ",
        "tue": "મંગળ",
        "wed": "બુધ",
        "thu": "ગુરુ",
        "fri": "શુક્ર",
        "sat": "શનિ",
        "sun": "રવિ",
    },
    "ur": {
        "app_title": "FaceAttend {ver}  ·  اسمارٹ حاضری نظام",
        "home_scanner": "ہوم اور اسکینر",
        "analytics": "تجزیات",
        "ai_buddy": "AI ساتھی",
        "messages": "پیغامات",
        "houses": "ہاؤس مقابلہ",
        "console": "کنسول",
        "settings": "ترتیبات",
        "start_scan": "▶  اسکین شروع کریں",
        "stop_scan": "■  اسکین روکیں",
        "refresh": "↺  ریفریش",
        "add_student": "➕  شامل کریں",
        "edit_student": "✏  ترمیم کریں",
        "remove_student": "🗑  ہٹائیں",
        "import": "📥  درآمد",
        "export": "⬇  برآمد",
        "total_students": "کل طلباء",
        "present_today": "آج حاضر",
        "absent_today": "آج غائب",
        "xp_leader": "XP لیڈر",
        "enrolled": "اندراج شدہ",
        "scanned_in": "اسکین ہوا",
        "not_yet_seen": "ابھی نہیں دیکھا",
        "top_scorer": "سب سے زیادہ سکور",
        "teacher_dashboard": "ٹیچر ڈیش بورڈ",
        "search": "طلباء تلاش کریں...",
        "clear": "صاف کریں",
        "last_recognition": "آخری شناخت",
        "waiting": "طالب علم کا انتظار…",
        "start_scanning": "حاضری ریکارڈ کرنے کے لیے اسکیننگ شروع کریں۔",
        "activity_feed": "سرگرمی فیڈ",
        "no_activity": "ابھی کوئی سرگرمی نہیں — طالب علم شامل کریں یا اسکیننگ شروع کریں۔",
        "camera_off": "📷  کیمرہ بند ہے",
        "settings_saved": "ترتیبات محفوظ ہوگئیں۔",
        "settings_reset": "ترتیبات ڈیفالٹ پر بحال۔",
        "backup_saved": "بیک اپ محفوظ ہوگیا۔",
        "email_sent": "{name} کے والدین کو ای میل بھیجی گئی",
        "whatsapp_opened": "{name} کے والدین کے لیے WhatsApp کھولا گیا",
        "student_added": "{name} کامیابی سے شامل ہوگیا!",
        "student_updated": "{name} اپڈیٹ ہوگیا۔",
        "student_removed": "{name} ہٹا دیا گیا۔",
        "select_student_first": "پہلے ایک طالب علم منتخب کریں۔",
        "attendance_late": "تاخیر سے",
        "attendance_ontime": "وقت پر",
        "welcome": "خوش آمدید، {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "زبان",
        "theme": "تھیم",
        "dark": "ڈارک",
        "light": "لائٹ",
        "schedule": "شیڈول",
        "badges": "بیجز اور کامیابیاں",
        "badge_first_steps": "پہلا قدم",
        "badge_week_warrior": "ویک واریئر",
        "badge_month_master": "ماہ ماسٹر",
        "badge_early_bird": "ارلی برڈ",
        "badge_centurion": "سینچورین",
        "badge_house_hero": "ہاؤس ہیرو",
        "qr_code": "QR کوڈ",
        "generate_qr": "QR بنائیں",
        "scan_qr": "QR اسکین کریں",
        "liveness": "لائیونیس کی شناخت",
        "voice_feedback": "صوتی تاثرات",
        "multi_camera": "ملٹی کیمرہ",
        "camera_select": "کیمرہ منتخب کریں",
        "api_server": "REST API سرور",
        "api_running": "پورٹ {port} پر API چل رہا ہے",
        "api_stopped": "API بند ہوگیا",
        "calendar": "حاضری کیلنڈر",
        "jan": "جنوری",
        "feb": "فروری",
        "mar": "مارچ",
        "apr": "اپریل",
        "may": "مئی",
        "jun": "جون",
        "mark_attendance": "حاضری درج کریں", "change_section": "سیکشن تبدیل کریں", "select_all": "سب منتخب کریں",
        "jul": "جولائی",
        "aug": "اگست",
        "sep": "ستمبر",
        "oct": "اکتوبر",
        "nov": "نومبر",
        "dec": "دسمبر",
        "whiteboard": "وائٹ بورڈ", "notepad": "نوٹ پیڈ",
        "mon": "پیر",
        "tue": "منگل",
        "wed": "بدھ",
        "thu": "جمعرات",
        "fri": "جمعہ",
        "sat": "ہفتہ",
        "sun": "اتوار",
    },
    "kn": {
        "app_title": "FaceAttend {ver}  ·  ಸ್ಮಾರ್ಟ್ ಹಾಜರಾತಿ ವ್ಯವಸ್ಥೆ",
        "home_scanner": "ಮುಖಪುಟ ಮತ್ತು ಸ್ಕ್ಯಾನರ್",
        "analytics": "ವಿಶ್ಲೇಷಣೆ",
        "ai_buddy": "AI ಸಹಾಯಕ",
        "messages": "ಸಂದೇಶಗಳು",
        "houses": "ಹೌಸ್ ಸ್ಪರ್ಧೆ",
        "console": "ಕನ್ಸೋಲ್",
        "settings": "ಸೆಟ್ಟಿಂಗ್‌ಗಳು",
        "start_scan": "▶  ಸ್ಕ್ಯಾನ್ ಪ್ರಾರಂಭಿಸಿ",
        "stop_scan": "■  ಸ್ಕ್ಯಾನ್ ನಿಲ್ಲಿಸಿ",
        "refresh": "↺  ರಿಫ್ರೆಶ್",
        "add_student": "➕  ಸೇರಿಸಿ",
        "edit_student": "✏  ಸಂಪಾದಿಸಿ",
        "remove_student": "🗑  ತೆಗೆದುಹಾಕಿ",
        "import": "📥  ಆಮದು",
        "export": "⬇  ರಫ್ತು",
        "total_students": "ಒಟ್ಟು ವಿದ್ಯಾರ್ಥಿಗಳು",
        "present_today": "ಇಂದು ಹಾಜರಿ",
        "absent_today": "ಇಂದು ಗೈರು",
        "xp_leader": "XP ಮುಂದಾಳು",
        "enrolled": "ದಾಖಲಾಗಿದೆ",
        "scanned_in": "ಸ್ಕ್ಯಾನ್ ಆಗಿದೆ",
        "not_yet_seen": "ಇನ್ನೂ ಕಂಡಿಲ್ಲ",
        "top_scorer": "ಅತ್ಯುತ್ತಮ ಸ್ಕೋರರ್",
        "teacher_dashboard": "ಶಿಕ್ಷಕ ಡ್ಯಾಶ್‌ಬೋರ್ಡ್",
        "search": "ವಿದ್ಯಾರ್ಥಿಗಳನ್ನು ಹುಡುಕಿ...",
        "clear": "ತೆರವುಗೊಳಿಸಿ",
        "last_recognition": "ಕೊನೆಯ ಗುರುತಿಸುವಿಕೆ",
        "waiting": "ವಿದ್ಯಾರ್ಥಿಗಾಗಿ ಕಾಯುತ್ತಿದೆ…",
        "start_scanning": "ಹಾಜರಾತಿ ದಾಖಲಿಸಲು ಸ್ಕ್ಯಾನಿಂಗ್ ಪ್ರಾರಂಭಿಸಿ.",
        "activity_feed": "ಚಟುವಟಿಕೆ ಫೀಡ್",
        "no_activity": "ಇನ್ನೂ ಚಟುವಟಿಕೆ ಇಲ್ಲ — ವಿದ್ಯಾರ್ಥಿಯನ್ನು ಸೇರಿಸಿ ಅಥವಾ ಸ್ಕ್ಯಾನಿಂಗ್ ಪ್ರಾರಂಭಿಸಿ.",
        "camera_off": "📷  ಕ್ಯಾಮೆರಾ ಆಫ್ ಆಗಿದೆ",
        "settings_saved": "ಸೆಟ್ಟಿಂಗ್‌ಗಳನ್ನು ಉಳಿಸಲಾಗಿದೆ.",
        "settings_reset": "ಸೆಟ್ಟಿಂಗ್‌ಗಳನ್ನು ಡೀಫಾಲ್ಟ್‌ಗೆ ಮರುಹೊಂದಿಸಲಾಗಿದೆ.",
        "backup_saved": "ಬ್ಯಾಕಪ್ ಉಳಿಸಲಾಗಿದೆ.",
        "email_sent": "{name} ರ ಪೋಷಕರಿಗೆ ಇಮೇಲ್ ಕಳುಹಿಸಲಾಗಿದೆ",
        "whatsapp_opened": "{name} ರ ಪೋಷಕರಿಗಾಗಿ WhatsApp ತೆರೆಯಲಾಗಿದೆ",
        "student_added": "{name} ಯಶಸ್ವಿಯಾಗಿ ಸೇರಿಸಲಾಗಿದೆ!",
        "student_updated": "{name} ಅಪ್ಡೇಟ್ ಆಗಿದೆ.",
        "student_removed": "{name} ತೆಗೆದುಹಾಕಲಾಗಿದೆ.",
        "select_student_first": "ಮೊದಲು ಒಬ್ಬ ವಿದ್ಯಾರ್ಥಿಯನ್ನು ಆಯ್ಕೆಮಾಡಿ.",
        "attendance_late": "ತಡವಾಗಿ",
        "attendance_ontime": "ಸಮಯಕ್ಕೆ",
        "welcome": "ಸ್ವಾಗತ, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "ಭಾಷೆ",
        "theme": "ಥೀಮ್",
        "dark": "ಡಾರ್ಕ್",
        "light": "ಲೈಟ್",
        "schedule": "ವೇಳಾಪಟ್ಟಿ",
        "badges": "ಬ್ಯಾಜ್‌ಗಳು ಮತ್ತು ಸಾಧನೆಗಳು",
        "badge_first_steps": "ಮೊದಲ ಹೆಜ್ಜೆಗಳು",
        "badge_week_warrior": "ವೀಕ್ ವಾರಿಯರ್",
        "badge_month_master": "ಮಂಥ್ ಮಾಸ್ಟರ್",
        "badge_early_bird": "ಅರ್ಲಿ ಬರ್ಡ್",
        "badge_centurion": "ಸೆಂಚುರಿಯನ್",
        "badge_house_hero": "ಹೌಸ್ ಹೀರೋ",
        "qr_code": "QR ಕೋಡ್",
        "generate_qr": "QR ರಚಿಸಿ",
        "scan_qr": "QR ಸ್ಕ್ಯಾನ್ ಮಾಡಿ",
        "liveness": "ಲೈವ್‌ನೆಸ್ ಪತ್ತೆ",
        "voice_feedback": "ಧ್ವನಿ ಪ್ರತಿಕ್ರಿಯೆ",
        "multi_camera": "ಮಲ್ಟಿ-ಕ್ಯಾಮೆರಾ",
        "camera_select": "ಕ್ಯಾಮೆರಾ ಆಯ್ಕೆಮಾಡಿ",
        "api_server": "REST API ಸರ್ವರ್",
        "api_running": "ಪೋರ್ಟ್ {port} ನಲ್ಲಿ API ಚಾಲನೆಯಲ್ಲಿದೆ",
        "api_stopped": "API ನಿಲ್ಲಿಸಲಾಗಿದೆ",
        "calendar": "ಹಾಜರಾತಿ ಕ್ಯಾಲೆಂಡರ್",
        "jan": "ಜನ",
        "feb": "ಫೆಬ್ರ",
        "mar": "ಮಾರ್ಚ್",
        "apr": "ಏಪ್ರಿ",
        "mark_attendance": "ಹಾಜರಾತಿ ದಾಖಲಿಸಿ", "change_section": "ವಿಭಾಗ ಬದಲಾಯಿಸಿ", "select_all": "ಎಲ್ಲಾ ಆಯ್ಕೆಮಾಡಿ",
        "may": "ಮೇ",
        "jun": "ಜೂನ್",
        "jul": "ಜುಲೈ",
        "aug": "ಆಗ",
        "sep": "ಸೆಪ್ಟೆಂ",
        "oct": "ಅಕ್ಟೋ",
        "nov": "ನವೆಂ",
        "whiteboard": "ವೈಟ್‌ಬೋರ್ಡ್", "notepad": "ನೋಟ್‌ಪ್ಯಾಡ್",
        "dec": "ಡಿಸೆಂ",
        "mon": "ಸೋಮ",
        "tue": "ಮಂಗಳ",
        "wed": "ಬುಧ",
        "thu": "ಗುರು",
        "fri": "ಶುಕ್ರ",
        "sat": "ಶನಿ",
        "sun": "ರವಿ",
    },
    "or": {
        "app_title": "FaceAttend {ver}  ·  ସ୍ମାର୍ଟ ଉପସ୍ଥିତି ସିଷ୍ଟମ",
        "home_scanner": "ହୋମ୍ ଓ ସ୍କାନର",
        "analytics": "ବିଶ୍ଳେଷଣ",
        "ai_buddy": "AI ସାଥୀ",
        "messages": "ବାର୍ତ୍ତା",
        "houses": "ହାଉସ୍ ପ୍ରତିଯୋଗିତା",
        "console": "କନସୋଲ୍",
        "settings": "ସେଟିଂସ୍",
        "start_scan": "▶  ସ୍କାନ୍ ଆରମ୍ଭ କରନ୍ତୁ",
        "stop_scan": "■  ସ୍କାନ୍ ବନ୍ଦ କରନ୍ତୁ",
        "refresh": "↺  ରିଫ୍ରେସ୍",
        "add_student": "➕  ଯୋଗ କରନ୍ତୁ",
        "edit_student": "✏  ସମ୍ପାଦନ କରନ୍ତୁ",
        "remove_student": "🗑  ହଟାନ୍ତୁ",
        "import": "📥  ଆମଦାନୀ",
        "export": "⬇  ରପ୍ତାନୀ",
        "total_students": "ସମୁଦାୟ ଛାତ୍ର",
        "present_today": "ଆଜି ଉପସ୍ଥିତ",
        "absent_today": "ଆଜି ଅନୁପସ୍ଥିତ",
        "xp_leader": "XP ନେତା",
        "enrolled": "ପଞ୍ଜିକୃତ",
        "scanned_in": "ସ୍କାନ ହୋଇଛି",
        "not_yet_seen": "ଏପର୍ଯ୍ୟନ୍ତ ଦେଖାଯାଇନାହିଁ",
        "top_scorer": "ସର୍ବୋଚ୍ଚ ସ୍କୋରର",
        "teacher_dashboard": "ଶିକ୍ଷକ ଡ୍ୟାସବୋର୍ଡ",
        "search": "ଛାତ୍ର ଖୋଜନ୍ତୁ...",
        "clear": "ସଫା କରନ୍ତୁ",
        "last_recognition": "ଅନ୍ତିମ ପରିଚୟ",
        "waiting": "ଛାତ୍ର ପାଇଁ ଅପେକ୍ଷା…",
        "start_scanning": "ଉପସ୍ଥିତି ଲେଖିବାକୁ ସ୍କାନିଂ ଆରମ୍ଭ କରନ୍ତୁ।",
        "activity_feed": "କାର୍ଯ୍ୟକଳାପ ଫିଡ୍",
        "no_activity": "ଏପର୍ଯ୍ୟନ୍ତ କୌଣସି କାର୍ଯ୍ୟକଳାପ ନାହିଁ — ଛାତ୍ର ଯୋଗ କରନ୍ତୁ କିମ୍ବା ସ୍କାନିଂ ଆରମ୍ଭ କରନ୍ତୁ।",
        "camera_off": "📷  କ୍ୟାମେରା ବନ୍ଦ ଅଛି",
        "settings_saved": "ସେଟିଂସ୍ ସେଭ ହୋଇଛି।",
        "settings_reset": "ସେଟିଂସ୍ ଡିଫଲ୍ଟକୁ ରିସେଟ ହୋଇଛି।",
        "backup_saved": "ବ୍ୟାକଅପ ସେଭ ହୋଇଛି।",
        "email_sent": "{name} ର ଅଭିଭାବକଙ୍କୁ ଇମେଲ ପଠାଗଲା",
        "whatsapp_opened": "{name} ର ଅଭିଭାବକ ପାଇଁ WhatsApp ଖୋଲାଗଲା",
        "student_added": "{name} ସଫଳତାର ସହିତ ଯୋଗ ହେଲା!",
        "student_updated": "{name} ଅପଡେଟ ହେଲା।",
        "student_removed": "{name} ହଟାଗଲା।",
        "select_student_first": "ପ୍ରଥମେ ଏକ ଛାତ୍ର ବାଛନ୍ତୁ।",
        "attendance_late": "ବିଳମ୍ବିତ",
        "attendance_ontime": "ସମୟରେ",
        "welcome": "ସ୍ୱାଗତ, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "ଭାଷା",
        "theme": "ଥିମ୍",
        "dark": "ଡାର୍କ",
        "light": "ଲାଇଟ୍",
        "schedule": "କାର୍ଯ୍ୟସୂଚୀ",
        "badges": "ବ୍ୟାଜ ଓ ସଫଳତା",
        "badge_first_steps": "ପ୍ରଥମ ପାଦବିକ୍ଷେପ",
        "badge_week_warrior": "ସପ୍ତାହ ଯୋଦ୍ଧା",
        "badge_month_master": "ମାସ ମାଷ୍ଟର",
        "badge_early_bird": "ଆର୍ଲି ବାର୍ଡ",
        "badge_centurion": "ସେଞ୍ଚୁରିଅନ୍",
        "badge_house_hero": "ହାଉସ୍ ହିରୋ",
        "qr_code": "QR କୋଡ୍",
        "generate_qr": "QR ତିଆରି କରନ୍ତୁ",
        "scan_qr": "QR ସ୍କାନ କରନ୍ତୁ",
        "liveness": "ଲାଇଭନେସ୍ ଚିହ୍ନଟ",
        "voice_feedback": "ସ୍ୱର ମତାମତ",
        "multi_camera": "ମଲ୍ଟି-କ୍ୟାମେରା",
        "camera_select": "କ୍ୟାମେରା ବାଛନ୍ତୁ",
        "api_server": "REST API ସର୍ଭର",
        "api_running": "ପୋର୍ଟ {port} ରେ API ଚାଲୁଅଛି",
        "api_stopped": "API ବନ୍ଦ ହୋଇଛି",
        "calendar": "ଉପସ୍ଥିତି କ୍ୟାଲେଣ୍ଡର",
        "jan": "ଜାନୁ",
        "feb": "ଫେବୃ",
        "mark_attendance": "ଉପସ୍ଥିତି ଦାଖଲ କରନ୍ତୁ", "change_section": "ବିଭାଗ ବଦଳାନ୍ତୁ", "select_all": "ସମସ୍ତ ବାଛନ୍ତୁ",
        "mar": "ମାର୍ଚ୍ଚ",
        "apr": "ଅପ୍ରେ",
        "may": "ମେ",
        "jun": "ଜୁନ",
        "jul": "ଜୁଲା",
        "aug": "ଅଗ",
        "sep": "ସେପ୍ଟେ",
        "oct": "ଅକ୍ଟୋ",
        "whiteboard": "ୱାଇଟବୋର୍ଡ", "notepad": "ନୋଟପେଡ",
        "nov": "ନଭେ",
        "dec": "ଡିସେ",
        "mon": "ସୋମ",
        "tue": "ମଙ୍ଗଳ",
        "wed": "ବୁଧ",
        "thu": "ଗୁରୁ",
        "fri": "ଶୁକ୍ର",
        "sat": "ଶନି",
        "sun": "ରବି",
    },
    "ml": {
        "app_title": "FaceAttend {ver}  ·  സ്മാർട്ട് ഹാജർ സിസ്റ്റം",
        "home_scanner": "ഹോം & സ്കാനർ",
        "analytics": "അനലിറ്റിക്സ്",
        "ai_buddy": "AI ബഡി",
        "messages": "സന്ദേശങ്ങൾ",
        "houses": "ഹൗസ് മത്സരം",
        "console": "കൺസോൾ",
        "settings": "സെറ്റിംഗ്സ്",
        "start_scan": "▶  സ്കാൻ ആരംഭിക്കുക",
        "stop_scan": "■  സ്കാൻ നിർത്തുക",
        "refresh": "↺  പുതുക്കുക",
        "add_student": "➕  ചേർക്കുക",
        "edit_student": "✏  എഡിറ്റ് ചെയ്യുക",
        "remove_student": "🗑  നീക്കം ചെയ്യുക",
        "import": "📥  ഇറക്കുമതി",
        "export": "⬇  കയറ്റുമതി",
        "total_students": "മൊത്തം വിദ്യാർത്ഥികൾ",
        "present_today": "ഇന്ന് ഹാജർ",
        "absent_today": "ഇന്ന് ഹാജരല്ല",
        "xp_leader": "XP ലീഡർ",
        "enrolled": "എൻറോൾ ചെയ്തു",
        "scanned_in": "സ്കാൻ ചെയ്തു",
        "not_yet_seen": "ഇതുവരെ കണ്ടിട്ടില്ല",
        "top_scorer": "ടോപ്പ് സ്കോററ്",
        "teacher_dashboard": "ടീച്ചർ ഡാഷ്ബോർഡ്",
        "search": "വിദ്യാർത്ഥികളെ തിരയുക...",
        "clear": "മായ്ക്കുക",
        "last_recognition": "അവസാന തിരിച്ചറിയൽ",
        "waiting": "വിദ്യാർത്ഥിക്കായി കാത്തിരിക്കുന്നു…",
        "start_scanning": "ഹാജർ രേഖപ്പെടുത്താൻ സ്കാനിംഗ് ആരംഭിക്കുക.",
        "activity_feed": "ആക്റ്റിവിറ്റി ഫീഡ്",
        "no_activity": "ഇതുവരെ പ്രവർത്തനമൊന്നുമില്ല — ഒരു വിദ്യാർത്ഥിയെ ചേർക്കുക അല്ലെങ്കിൽ സ്കാനിംഗ് ആരംഭിക്കുക.",
        "camera_off": "📷  ക്യാമറ ഓഫാണ്",
        "settings_saved": "സെറ്റിംഗ്സ് സംരക്ഷിച്ചു.",
        "settings_reset": "സെറ്റിംഗ്സ് ഡിഫോൾട്ടിലേക്ക് പുനഃസജ്ജമാക്കി.",
        "backup_saved": "ബാക്കപ്പ് സംരക്ഷിച്ചു.",
        "email_sent": "{name} ന്റെ രക്ഷിതാവിന് ഇമെയിൽ അയച്ചു",
        "whatsapp_opened": "{name} ന്റെ രക്ഷിതാവിനായി WhatsApp തുറന്നു",
        "student_added": "{name} വിജയകരമായി ചേർത്തു!",
        "student_updated": "{name} അപ്ഡേറ്റ് ചെയ്തു.",
        "student_removed": "{name} നീക്കം ചെയ്തു.",
        "select_student_first": "ആദ്യം ഒരു വിദ്യാർത്ഥിയെ തിരഞ്ഞെടുക്കുക.",
        "attendance_late": "വൈകി",
        "attendance_ontime": "സമയത്ത്",
        "welcome": "സ്വാഗതം, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "ഭാഷ",
        "theme": "തീം",
        "dark": "ഡാർക്ക്",
        "light": "ലൈറ്റ്",
        "schedule": "ഷെഡ്യൂൾ",
        "badges": "ബാഡ്ജുകളും നേട്ടങ്ങളും",
        "badge_first_steps": "ആദ്യ ചുവടുകൾ",
        "badge_week_warrior": "വീക്ക് വാരിയർ",
        "badge_month_master": "മന്ത് മാസ്റ്റർ",
        "badge_early_bird": "ഏർലി ബേർഡ്",
        "badge_centurion": "സെഞ്ച്വറിയൻ",
        "badge_house_hero": "ഹൗസ് ഹീറോ",
        "qr_code": "QR കോഡ്",
        "generate_qr": "QR ഉണ്ടാക്കുക",
        "scan_qr": "QR സ്കാൻ ചെയ്യുക",
        "liveness": "ലൈവ്നെസ് കണ്ടെത്തൽ",
        "voice_feedback": "വോയിസ് ഫീഡ്ബാക്ക്",
        "multi_camera": "മൾട്ടി-ക്യാമറ",
        "camera_select": "ക്യാമറ തിരഞ്ഞെടുക്കുക",
        "api_server": "REST API സെർവർ",
        "api_running": "പോർട്ട് {port} ൽ API പ്രവർത്തിക്കുന്നു",
        "api_stopped": "API നിർത്തി",
        "calendar": "ഹാജർ കലണ്ടർ",
        "mark_attendance": "ഹാജർ രേഖപ്പെടുത്തുക", "change_section": "വിഭാഗം മാറ്റുക", "select_all": "എല്ലാം തിരഞ്ഞെടുക്കുക",
        "jan": "ജനു",
        "feb": "ഫെബ്ര",
        "mar": "മാർച്ച്",
        "apr": "ഏപ്രി",
        "may": "മേയ്",
        "jun": "ജൂൺ",
        "jul": "ജൂലൈ",
        "aug": "ഓഗ",
        "sep": "സെപ്റ്റ",
        "whiteboard": "വൈറ്റ്ബോർഡ്", "notepad": "നോട്ട്പാഡ്",
        "oct": "ഒക്ടോ",
        "nov": "നവം",
        "dec": "ഡിസം",
        "mon": "തിങ്കൾ",
        "tue": "ചൊവ്വ",
        "wed": "ബുധൻ",
        "thu": "വ്യാഴം",
        "fri": "വെള്ളി",
        "sat": "ശനി",
        "sun": "ഞായർ",
    },
    "pa": {
        "app_title": "FaceAttend {ver}  ·  ਸਮਾਰਟ ਹਾਜ਼ਰੀ ਸਿਸਟਮ",
        "home_scanner": "ਹੋਮ ਅਤੇ ਸਕੈਨਰ",
        "analytics": "ਵਿਸ਼ਲੇਸ਼ਣ",
        "ai_buddy": "AI ਸਾਥੀ",
        "messages": "ਸੰਦੇਸ਼",
        "houses": "ਹਾਊਸ ਮੁਕਾਬਲਾ",
        "console": "ਕੰਸੋਲ",
        "settings": "ਸੈਟਿੰਗਾਂ",
        "start_scan": "▶  ਸਕੈਨ ਸ਼ੁਰੂ ਕਰੋ",
        "stop_scan": "■  ਸਕੈਨ ਰੋਕੋ",
        "refresh": "↺  ਤਾਜ਼ਾ ਕਰੋ",
        "add_student": "➕  ਸ਼ਾਮਲ ਕਰੋ",
        "edit_student": "✏  ਸੰਪਾਦਿਤ ਕਰੋ",
        "remove_student": "🗑  ਹਟਾਓ",
        "import": "📥  ਆਯਾਤ",
        "export": "⬇  ਨਿਰਯਾਤ",
        "total_students": "ਕੁੱਲ ਵਿਦਿਆਰਥੀ",
        "present_today": "ਅੱਜ ਹਾਜ਼ਰ",
        "absent_today": "ਅੱਜ ਗੈਰਹਾਜ਼ਰ",
        "xp_leader": "XP ਲੀਡਰ",
        "enrolled": "ਦਰਜ ਕੀਤਾ",
        "scanned_in": "ਸਕੈਨ ਕੀਤਾ",
        "not_yet_seen": "ਅਜੇ ਨਹੀਂ ਦੇਖਿਆ",
        "top_scorer": "ਸਭ ਤੋਂ ਵੱਧ ਸਕੋਰਰ",
        "teacher_dashboard": "ਟੀਚਰ ਡੈਸ਼ਬੋਰਡ",
        "search": "ਵਿਦਿਆਰਥੀ ਲੱਭੋ...",
        "clear": "ਸਾਫ਼ ਕਰੋ",
        "last_recognition": "ਆਖਰੀ ਪਛਾਣ",
        "waiting": "ਵਿਦਿਆਰਥੀ ਦੀ ਉਡੀਕ…",
        "start_scanning": "ਹਾਜ਼ਰੀ ਦਰਜ ਕਰਨ ਲਈ ਸਕੈਨਿੰਗ ਸ਼ੁਰੂ ਕਰੋ।",
        "activity_feed": "ਸਰਗਰਮੀ ਫੀਡ",
        "no_activity": "ਅਜੇ ਕੋਈ ਸਰਗਰਮੀ ਨਹੀਂ — ਵਿਦਿਆਰਥੀ ਸ਼ਾਮਲ ਕਰੋ ਜਾਂ ਸਕੈਨਿੰਗ ਸ਼ੁਰੂ ਕਰੋ।",
        "camera_off": "📷  ਕੈਮਰਾ ਬੰਦ ਹੈ",
        "settings_saved": "ਸੈਟਿੰਗਾਂ ਸੇਵ ਹੋ ਗਈਆਂ।",
        "settings_reset": "ਸੈਟਿੰਗਾਂ ਡਿਫਾਲਟ ਤੇ ਰੀਸੈਟ ਹੋਈਆਂ।",
        "backup_saved": "ਬੈਕਅੱਪ ਸੇਵ ਹੋਇਆ।",
        "email_sent": "{name} ਦੇ ਮਾਪਿਆਂ ਨੂੰ ਈਮੇਲ ਭੇਜੀ ਗਈ",
        "whatsapp_opened": "{name} ਦੇ ਮਾਪਿਆਂ ਲਈ WhatsApp ਖੋਲ੍ਹਿਆ ਗਿਆ",
        "student_added": "{name} ਸਫਲਤਾਪੂਰਵਕ ਸ਼ਾਮਲ ਹੋਇਆ!",
        "student_updated": "{name} ਅੱਪਡੇਟ ਹੋਇਆ।",
        "student_removed": "{name} ਹਟਾਇਆ ਗਿਆ।",
        "select_student_first": "ਪਹਿਲਾਂ ਇੱਕ ਵਿਦਿਆਰਥੀ ਚੁਣੋ।",
        "attendance_late": "ਦੇਰ ਨਾਲ",
        "attendance_ontime": "ਸਮੇਂ ਸਿਰ",
        "welcome": "ਜੀ ਆਇਆਂ ਨੂੰ, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "ਭਾਸ਼ਾ",
        "theme": "ਥੀਮ",
        "dark": "ਡਾਰਕ",
        "light": "ਲਾਈਟ",
        "schedule": "ਸਮਾਂ-ਸਾਰਣੀ",
        "badges": "ਬੈਜ ਅਤੇ ਪ੍ਰਾਪਤੀਆਂ",
        "badge_first_steps": "ਪਹਿਲੇ ਕਦਮ",
        "badge_week_warrior": "ਵੀਕ ਵਾਰੀਅਰ",
        "badge_month_master": "ਮਹੀਨਾ ਮਾਸਟਰ",
        "badge_early_bird": "ਅਰਲੀ ਬਰਡ",
        "badge_centurion": "ਸੈਂਚੁਰੀਅਨ",
        "badge_house_hero": "ਹਾਊਸ ਹੀਰੋ",
        "qr_code": "QR ਕੋਡ",
        "generate_qr": "QR ਬਣਾਓ",
        "scan_qr": "QR ਸਕੈਨ ਕਰੋ",
        "liveness": "ਲਾਈਵਨੈੱਸ ਖੋਜ",
        "voice_feedback": "ਆਵਾਜ਼ ਫੀਡਬੈਕ",
        "multi_camera": "ਮਲਟੀ-ਕੈਮਰਾ",
        "camera_select": "ਕੈਮਰਾ ਚੁਣੋ",
        "api_server": "REST API ਸਰਵਰ",
        "api_running": "ਪੋਰਟ {port} ਤੇ API ਚੱਲ ਰਿਹਾ ਹੈ",
        "mark_attendance": "ਹਾਜ਼ਰੀ ਦਰਜ ਕਰੋ", "change_section": "ਸੈਕਸ਼ਨ ਬਦਲੋ", "select_all": "ਸਾਰੇ ਚੁਣੋ",
        "api_stopped": "API ਬੰਦ ਹੋ ਗਿਆ",
        "calendar": "ਹਾਜ਼ਰੀ ਕੈਲੰਡਰ",
        "jan": "ਜਨ",
        "feb": "ਫਰ",
        "mar": "ਮਾਰ",
        "apr": "ਅਪ੍ਰੈ",
        "may": "ਮਈ",
        "jun": "ਜੂਨ",
        "jul": "ਜੁਲ",
        "aug": "ਅਗ",
        "whiteboard": "ਵ੍ਹਾਈਟਬੋਰਡ", "notepad": "ਨੋਟਪੈਡ",
        "sep": "ਸਤੰ",
        "oct": "ਅਕਤੂ",
        "nov": "ਨਵੰ",
        "dec": "ਦਸੰ",
        "mon": "ਸੋਮ",
        "tue": "ਮੰਗਲ",
        "wed": "ਬੁੱਧ",
        "thu": "ਵੀਰ",
        "fri": "ਸ਼ੁੱਕਰ",
        "sat": "ਸ਼ਨੀ",
        "sun": "ਐਤ",
    },
    "as": {
        "app_title": "FaceAttend {ver}  ·  স্মাৰ্ট উপস্থিতি প্ৰণালী",
        "home_scanner": "হোম আৰু স্কেনাৰ",
        "analytics": "বিশ্লেষণ",
        "ai_buddy": "AI সহায়ক",
        "messages": "বাৰ্তা",
        "houses": "হাউস প্ৰতিযোগিতা",
        "console": "কনসোল",
        "settings": "ছেটিংছ",
        "start_scan": "▶  স্কেন আৰম্ভ কৰক",
        "stop_scan": "■  স্কেন বন্ধ কৰক",
        "refresh": "↺  ৰিফ্ৰেশ",
        "add_student": "➕  যোগ কৰক",
        "edit_student": "✏  সম্পাদনা কৰক",
        "remove_student": "🗑  আঁতৰাওক",
        "import": "📥  আমদানি",
        "export": "⬇  ৰপ্তানি",
        "total_students": "মুঠ শিক্ষাৰ্থী",
        "present_today": "আজি উপস্থিত",
        "absent_today": "আজি অনুপস্থিত",
        "xp_leader": "XP নেতা",
        "enrolled": "নাম লগোয়া হৈছে",
        "scanned_in": "স্কেন কৰা হৈছে",
        "not_yet_seen": "এতিয়াও দেখা যায়নি",
        "top_scorer": "সৰ্বোচ্চ স্কোৰাৰ",
        "teacher_dashboard": "শিক্ষক ডেশবোৰ্ড",
        "search": "শিক্ষাৰ্থী বিচাৰক...",
        "clear": "পৰিষ্কাৰ কৰক",
        "last_recognition": "সৰ্বশেষ স্বীকৃতি",
        "waiting": "শিক্ষাৰ্থীৰ বাবে অপেক্ষা…",
        "start_scanning": "উপস্থিতি লিপিবদ্ধ কৰিবলে স্কেনিং আৰম্ভ কৰক।",
        "activity_feed": "কাৰ্যকলাপ ফিড",
        "no_activity": "এতিয়াও কোনো কাৰ্যকলাপ নাই — শিক্ষাৰ্থী যোগ কৰক অথবা স্কেনিং আৰম্ভ কৰক।",
        "camera_off": "📷  কেমেৰা বন্ধ আছে",
        "settings_saved": "ছেটিংছ সংৰক্ষণ কৰা হৈছে।",
        "settings_reset": "ছেটিংছ ডিফল্টলে ৰিসেট কৰা হৈছে।",
        "backup_saved": "বেকআপ সংৰক্ষণ কৰা হৈছে।",
        "email_sent": "{name} ৰ অভিভাৱকলে ইমেইল পঠিওৱা হৈছে",
        "whatsapp_opened": "{name} ৰ অভিভাৱকৰ বাবে WhatsApp খোলা হৈছে",
        "student_added": "{name} সফলতাৰে যোগ কৰা হৈছে!",
        "student_updated": "{name} আপডেট কৰা হৈছে।",
        "student_removed": "{name} আঁতৰোৱা হৈছে।",
        "select_student_first": "প্ৰথমে এজন শিক্ষাৰ্থী বাছনি কৰক।",
        "attendance_late": "পলম",
        "attendance_ontime": "সময়মতে",
        "welcome": "স্বাগতম, {name}",
        "marked_status": "{status}  ·  {roll}  ·  +{xp} XP  ·  {time}",
        "language": "ভাষা",
        "theme": "থীম",
        "dark": "ডাৰ্ক",
        "light": "লাইট",
        "schedule": "সময়সূচী",
        "badges": "বেজ আৰু সফলতা",
        "badge_first_steps": "প্ৰথম খোজ",
        "badge_week_warrior": "সপ্তাহৰ যোদ্ধা",
        "badge_month_master": "মাহৰ মাষ্টাৰ",
        "badge_early_bird": "আৰ্লি বাৰ্ড",
        "badge_centurion": "চেঞ্চুৰিয়ন",
        "badge_house_hero": "হাউস হিৰো",
        "qr_code": "QR কোড",
        "generate_qr": "QR বনাওক",
        "scan_qr": "QR স্কেন কৰক",
        "liveness": "লাইভনেছ চিনাক্তকৰণ",
        "voice_feedback": "মাত প্ৰতিক্ৰিয়া",
        "multi_camera": "মাল্টি-কেমেৰা",
        "camera_select": "কেমেৰা বাছনি কৰক",
        "mark_attendance": "উপস্থিতি লিখক", "change_section": "শাখা সলনি কৰক", "select_all": "সকলো বাছক",
        "api_server": "REST API চাৰ্ভাৰ",
        "api_running": "পোৰ্ট {port} ত API চলি আছে",
        "api_stopped": "API বন্ধ হৈছে",
        "calendar": "উপস্থিতি কেলেণ্ডাৰ",
        "jan": "জানু",
        "feb": "ফেব্ৰু",
        "mar": "মাৰ্চ",
        "apr": "এপ্ৰি",
        "may": "মে",
        "jun": "জুন",
        "jul": "জুলাই",
        "whiteboard": "ৱাইটবৰ্ড", "notepad": "নোটপেড",
        "aug": "আগ",
        "sep": "সেপ্তে",
        "oct": "অক্টো",
        "nov": "নৱে",
        "dec": "ডিসে",
        "mon": "সোম",
        "tue": "মঙ্গল",
        "wed": "বুধ",
        "thu": "বৃহ",
        "fri": "শুক্ৰ",
        "sat": "শনি",
        "sun": "ৰবি",
    },
}

LANG_NAMES = {
    "en": "English",
    "hi": "हिन्दी (Hindi)",
    "bn": "বাংলা (Bengali)",
    "te": "తెలుగు (Telugu)",
    "mr": "मराठी (Marathi)",
    "ta": "தமிழ் (Tamil)",
    "gu": "ગુજરાતી (Gujarati)",
    "ur": "اردو (Urdu)",
    "kn": "ಕನ್ನಡ (Kannada)",
    "or": "ଓଡ଼ିଆ (Odia)",
    "ml": "മലയാളം (Malayalam)",
    "pa": "ਪੰਜਾਬੀ (Punjabi)",
    "as": "অসমীয়া (Assamese)",
}


def _(key: str, lang: str = "en", **kwargs) -> str:
    """Translate a key into the given language. Falls back to English."""
    text = TRANSLATIONS.get(lang, TRANSLATIONS["en"]).get(key, TRANSLATIONS["en"].get(key, key))
    try:
        return text.format(**kwargs)
    except Exception:
        return text


# ═══════════════════════════════════════════════════════════════════════════════
#  THEME SYSTEM  (Light / Dark toggle)
# ═══════════════════════════════════════════════════════════════════════════════
class Theme:
    """Dynamic palette that swaps between Dark and Light modes."""

    DARK = type("P", (), {
        "rail": "#0a0e15", "bg": "#0e1117", "sidebar": "#12161f",
        "sidebar2": "#1a2130", "surface": "#1c2233", "surface2": "#222b3a",
        "hover": "#283248", "border": "#2a3448", "border_soft": "#202838",
        "text": "#e8edf5", "muted": "#7b8aa0", "faint": "#4d586c",
        "blue": "#3b82f6", "blue_dark": "#2563eb", "blue_glow": "#1d4ed8",
        "blue_soft": "#132236", "green": "#22c55e", "green_dark": "#16a34a",
        "green_soft": "#0d2218", "amber": "#f59e0b", "amber_soft": "#241b08",
        "red": "#ef4444", "red_dark": "#dc2626", "red_soft": "#2a1316",
        "purple": "#a855f7", "purple_soft": "#1c1430", "ink": "#f1f5f9",
        "tint_blue": "#0f1e35", "tint_green": "#0d2218",
        "tint_amber": "#231a05", "tint_purple": "#1a1030",
        "chart_bg": "#0e1117", "chart_fg": "#e8edf5",
        "chart_grid": "#2a3448", "calendar_present": "#22c55e",
        "calendar_absent": "#ef4444", "calendar_neutral": "#2a3448",
    })()

    LIGHT = type("P", (), {
        "rail": "#f1f5f9", "bg": "#ffffff", "sidebar": "#f8fafc",
        "sidebar2": "#f1f5f9", "surface": "#ffffff", "surface2": "#f8fafc",
        "hover": "#e2e8f0", "border": "#cbd5e1", "border_soft": "#e2e8f0",
        "text": "#0f172a", "muted": "#64748b", "faint": "#94a3b8",
        "blue": "#2563eb", "blue_dark": "#1d4ed8", "blue_glow": "#1e40af",
        "blue_soft": "#dbeafe", "green": "#16a34a", "green_dark": "#15803d",
        "green_soft": "#dcfce7", "amber": "#d97706", "amber_soft": "#fef3c7",
        "red": "#dc2626", "red_dark": "#b91c1c", "red_soft": "#fee2e2",
        "purple": "#7c3aed", "purple_soft": "#ede9fe", "ink": "#0f172a",
        "tint_blue": "#eff6ff", "tint_green": "#f0fdf4",
        "tint_amber": "#fffbeb", "tint_purple": "#f5f3ff",
        "chart_bg": "#ffffff", "chart_fg": "#0f172a",
        "chart_grid": "#e2e8f0", "calendar_present": "#16a34a",
        "calendar_absent": "#dc2626", "calendar_neutral": "#e2e8f0",
    })()

    @classmethod
    def get(cls, is_dark: bool = True):
        return cls.DARK if is_dark else cls.LIGHT


# ═══════════════════════════════════════════════════════════════════════════════
#  SCHEDULE MANAGER
# ═══════════════════════════════════════════════════════════════════════════════
SCHEDULE_FILE = DATA_DIR / "schedule.json"

class ScheduleManager:
    """Per-day late thresholds and active/inactive day configuration."""

    DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    def __init__(self):
        self._data: dict = {}
        self.load()

    def load(self):
        if SCHEDULE_FILE.exists():
            try:
                with SCHEDULE_FILE.open() as f:
                    self._data = json.load(f)
            except Exception:
                self._data = {}
        else:
            self._reset_defaults()

    def save(self):
        with SCHEDULE_FILE.open("w") as f:
            json.dump(self._data, f, indent=2)

    def _reset_defaults(self):
        self._data = {
            day: {"active": day not in ("Sat", "Sun"), "late_time": "08:32"}
            for day in self.DAYS
        }
        self.save()

    def is_active(self, day_name: str | None = None) -> bool:
        if day_name is None:
            day_name = datetime.now().strftime("%a")
        return self._data.get(day_name, {}).get("active", True)

    def late_time(self, day_name: str | None = None) -> dtime:
        if day_name is None:
            day_name = datetime.now().strftime("%a")
        raw = self._data.get(day_name, {}).get("late_time", "08:32")
        try:
            h, m = raw.split(":")
            return dtime(int(h), int(m))
        except Exception:
            return dtime(8, 32)

    def set_day(self, day: str, active: bool, late_time: str):
        if day not in self._data:
            self._data[day] = {}
        self._data[day]["active"] = active
        self._data[day]["late_time"] = late_time
        self.save()


# ═══════════════════════════════════════════════════════════════════════════════
#  BADGE & ACHIEVEMENT SYSTEM
# ═══════════════════════════════════════════════════════════════════════════════
BADGE_FILE = DATA_DIR / "badges.json"

class BadgeSystem:
    """Auto-awards attendance-based badges to students."""

    BADGE_DEFS = {
        "first_steps":  {"name": "badge_first_steps",  "icon": "👣", "req_days": 1},
        "week_warrior": {"name": "badge_week_warrior", "icon": "⚔", "req_days": 5},
        "month_master": {"name": "badge_month_master", "icon": "👑", "req_days": 20},
        "early_bird":   {"name": "badge_early_bird",   "icon": "🐦", "req_early": 10},
        "centurion":    {"name": "badge_centurion",    "icon": "🏛", "req_days": 100},
        "house_hero":   {"name": "badge_house_hero",   "icon": "🏆", "req_house_pct": 90.0},
    }

    def __init__(self, store: AttendanceStore):
        self.store = store
        self._awards: dict[str, list[str]] = {}
        self.load()

    def load(self):
        if BADGE_FILE.exists():
            try:
                with BADGE_FILE.open() as f:
                    self._awards = json.load(f)
            except Exception:
                self._awards = {}
        else:
            self._awards = {}

    def save(self):
        with BADGE_FILE.open("w") as f:
            json.dump(self._awards, f, indent=2)

    def get_badges(self, uid: str) -> list[dict]:
        """Return list of badge dicts for a student."""
        earned = self._awards.get(uid, [])
        return [b for k, b in self.BADGE_DEFS.items() if k in earned]

    def check_and_award(self, uid: str, student: dict, is_early: bool = False):
        """Check if student qualifies for new badges and award them."""
        if uid not in self._awards:
            self._awards[uid] = []
        earned = self._awards[uid]
        days = safe_int(student.get("No. of Days Present", 0))
        new = []

        for key, badge in self.BADGE_DEFS.items():
            if key in earned:
                continue
            if "req_days" in badge and days >= badge["req_days"]:
                earned.append(key); new.append(badge)
            elif "req_early" in badge and is_early:
                # Count early arrivals from audit log (simplified)
                early_count = self._count_early_arrivals(uid)
                if early_count >= badge["req_early"]:
                    earned.append(key); new.append(badge)
            elif "req_house_pct" in badge:
                house = str(student.get("House", "None")).strip()
                if house != "None":
                    stats = self.store.house_stats()
                    if stats.get(house, {}).get("pct", 0) >= badge["req_house_pct"]:
                        earned.append(key); new.append(badge)

        if new:
            self.save()
        return new

    def _count_early_arrivals(self, uid: str) -> int:
        rows = self.store._load_audit_rows()
        count = 0
        for row in rows:
            if len(row) > 7 and row[2] == uid and row[7] == "On-time":
                count += 1
        return count


# ═══════════════════════════════════════════════════════════════════════════════
#  AES-256 ENCRYPTION MANAGER
# ═══════════════════════════════════════════════════════════════════════════════
KEY_FILE = DATA_DIR / ".enc_key"

class EncryptionManager:
    """Fernet symmetric encryption for sensitive parent contact fields."""

    def __init__(self):
        self._fernet = None
        if Fernet is not None:
            self._init_key()

    def _init_key(self):
        if KEY_FILE.exists():
            try:
                with KEY_FILE.open("rb") as f:
                    key = f.read()
                self._fernet = Fernet(key)
            except Exception:
                self._generate_key()
        else:
            self._generate_key()

    def _generate_key(self):
        key = Fernet.generate_key()
        with KEY_FILE.open("wb") as f:
            f.write(key)
        self._fernet = Fernet(key)

    def encrypt(self, text: str) -> str:
        if self._fernet is None or not text:
            return text
        try:
            return self._fernet.encrypt(text.encode()).decode()
        except Exception:
            return text

    def decrypt(self, token: str) -> str:
        if self._fernet is None or not token:
            return token
        try:
            return self._fernet.decrypt(token.encode()).decode()
        except Exception:
            return token

    def is_encrypted(self, text: str) -> bool:
        if not text or self._fernet is None:
            return False
        try:
            self._fernet.decrypt(text.encode())
            return True
        except Exception:
            return False


# ═══════════════════════════════════════════════════════════════════════════════
#  QR CODE SYSTEM
# ═══════════════════════════════════════════════════════════════════════════════
QR_DIR = DATA_DIR / "qr_codes"

def generate_student_qr(uid: str, name: str, roll: str) -> Path | None:
    """Generate a QR code PNG for a student. Returns path or None."""
    if qrcode is None:
        return None
    QR_DIR.mkdir(parents=True, exist_ok=True)
    path = QR_DIR / f"{uid}.png"
    data = json.dumps({"uid": uid, "name": name, "roll": roll, "app": "FaceAttend"})
    try:
        img = qrcode.make(data, box_size=6, border=2)
        img.save(str(path))
        return path
    except Exception:
        return None


def decode_qr_from_frame(frame) -> dict | None:
    """Try to decode a QR code from an OpenCV frame."""
    if cv2 is None:
        return None
    try:
        detector = cv2.QRCodeDetector()
        data, bbox, _ = detector.detectAndDecode(frame)
        if data:
            try:
                return json.loads(data)
            except Exception:
                return {"raw": data}
    except Exception:
        pass
    return None


# ═══════════════════════════════════════════════════════════════════════════════
#  LIVENESS DETECTION (Blink via EAR)
# ═══════════════════════════════════════════════════════════════════════════════
# dlib 68-point facial landmark indices
LEFT_EYE  = list(range(36, 42))
RIGHT_EYE = list(range(42, 48))

def eye_aspect_ratio(eye):
    """Compute EAR for a single eye given 6 (x,y) points."""
    if sci_dist is None:
        return 1.0
    A = sci_dist.euclidean(eye[1], eye[5])
    B = sci_dist.euclidean(eye[2], eye[4])
    C = sci_dist.euclidean(eye[0], eye[3])
    return (A + B) / (2.0 * C) if C > 0 else 1.0


class LivenessDetector:
    """Simple blink-based liveness detection."""

    EAR_THRESHOLD = 0.22
    CONSEC_FRAMES = 2

    def __init__(self):
        self._predictor = None
        self._detector = None
        self._counter = 0
        self._total_blinks = 0
        if dlib is not None:
            self._init_model()

    def _init_model(self):
        # Look for shape_predictor_68_face_landmarks.dat in assets or data
        model_paths = [
            ASSETS_DIR / "shape_predictor_68_face_landmarks.dat",
            DATA_DIR / "shape_predictor_68_face_landmarks.dat",
        ]
        for p in model_paths:
            if p.exists():
                try:
                    self._detector = dlib.get_frontal_face_detector()
                    self._predictor = dlib.shape_predictor(str(p))
                    break
                except Exception:
                    pass

    def is_live(self, frame) -> bool:
        """Return True if a blink is detected (indicating a real person)."""
        if self._predictor is None or self._detector is None:
            return True  # fallback: assume live if model unavailable
        try:
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            rects = self._detector(gray, 0)
            for rect in rects:
                shape = self._predictor(gray, rect)
                coords = [(shape.part(i).x, shape.part(i).y) for i in range(68)]
                left_ear = eye_aspect_ratio([coords[i] for i in LEFT_EYE])
                right_ear = eye_aspect_ratio([coords[i] for i in RIGHT_EYE])
                ear = (left_ear + right_ear) / 2.0
                if ear < self.EAR_THRESHOLD:
                    self._counter += 1
                else:
                    if self._counter >= self.CONSEC_FRAMES:
                        self._total_blinks += 1
                    self._counter = 0
                # Consider "live" if at least one blink has ever been detected
                # OR if eyes are currently open (EAR above threshold)
                return self._total_blinks > 0 or ear > self.EAR_THRESHOLD
        except Exception:
            pass
        return True

    def reset(self):
        self._counter = 0
        self._total_blinks = 0


# ═══════════════════════════════════════════════════════════════════════════════
#  VOICE FEEDBACK
# ═══════════════════════════════════════════════════════════════════════════════
_voice_engine = None

def get_voice_engine():
    global _voice_engine
    if _voice_engine is None and pyttsx3 is not None:
        try:
            _voice_engine = pyttsx3.init()
            _voice_engine.setProperty("rate", 170)
        except Exception:
            _voice_engine = None
    return _voice_engine


def speak_text(text: str):
    """Non-blocking text-to-speech announcement. (Voice disabled in 2.9)"""
    return  # voice removed in 2.9

    engine = get_voice_engine()
    if engine is None or not text:
        return

    def _talk():
        try:
            engine.say(text)
            engine.runAndWait()
        except Exception:
            pass

    threading.Thread(target=_talk, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════════════
#  MULTI-CAMERA DETECTION
# ═══════════════════════════════════════════════════════════════════════════════
def list_cameras(max_test: int = 3) -> list[int]:
    """Test camera indices 0..max_test and return those that open.
    Suppresses OpenCV C++ DSHOW spam by redirecting fd 2 at OS level."""
    if cv2 is None:
        return [0]
    import sys, io
    found = []
    # Redirect fd 2 (C-level stderr) so OpenCV DSHOW warnings don't appear
    devnull_fd = os.open(os.devnull, os.O_WRONLY)
    old_fd = os.dup(2)
    os.dup2(devnull_fd, 2)
    os.close(devnull_fd)
    try:
        for i in range(max_test):
            cap = cv2.VideoCapture(i, cv2.CAP_DSHOW) if os.name == "nt" else cv2.VideoCapture(i)
            if cap is not None and cap.isOpened():
                found.append(i)
                cap.release()
            elif cap is not None:
                cap.release()
    finally:
        os.dup2(old_fd, 2)
        os.close(old_fd)
    return found if found else [0]


# ═══════════════════════════════════════════════════════════════════════════════
#  REST API SERVER
# ═══════════════════════════════════════════════════════════════════════════════
API_PORT = 8765

class FaceAttendAPIHandler(BaseHTTPRequestHandler):
    """Minimal REST API for external integrations."""

    app_instance = None  # set by FaceAttendApp

    def log_message(self, format, *args):
        pass  # silence console spam

    def _send_json(self, status: int, data: dict):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        app = self.app_instance

        if path == "/api/students":
            if app is None:
                self._send_json(503, {"error": "App not ready"})
                return
            students = app.store.load_students()
            self._send_json(200, {"count": len(students), "students": students})

        elif path == "/api/today":
            if app is None:
                self._send_json(503, {"error": "App not ready"})
                return
            present = [app.students_by_uid.get(u, {}) for u in app.scanned_today]
            self._send_json(200, {
                "date": date.today().isoformat(),
                "present_count": len(present),
                "present": present,
            })

        elif path == "/api/houses":
            if app is None:
                self._send_json(503, {"error": "App not ready"})
                return
            stats = app.store.house_stats()
            self._send_json(200, stats)

        elif path == "/api/qr":
            params = urllib.parse.parse_qs(parsed.query)
            uid = params.get("uid", [""])[0]
            if not uid:
                self._send_json(400, {"error": "Missing uid param"})
                return
            path = QR_DIR / f"{uid}.png"
            if not path.exists():
                self._send_json(404, {"error": "QR not found"})
                return
            self.send_response(200)
            self.send_header("Content-Type", "image/png")
            self.end_headers()
            with open(path, "rb") as f:
                self.wfile.write(f.read())

        elif path == "/api/health":
            self._send_json(200, {"status": "ok", "version": APP_VERSION})

        else:
            self._send_json(404, {"error": "Unknown endpoint"})

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.end_headers()


class APIServer:
    """Threaded wrapper for the HTTP API."""

    def __init__(self, app):
        self.app = app
        self._server = None
        self._thread = None
        self._running = False

    def start(self):
        if HTTPServer is None or socketserver is None:
            return False
        try:
            FaceAttendAPIHandler.app_instance = self.app
            self._server = HTTPServer(("0.0.0.0", API_PORT), FaceAttendAPIHandler)
            self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
            self._thread.start()
            self._running = True
            return True
        except Exception as e:
            logger.error(f"API server failed: {e}")
            return False

    def stop(self):
        if self._server:
            try:
                self._server.shutdown()
            except Exception:
                pass
        self._running = False


# ═══════════════════════════════════════════════════════════════════════════════
#  CALENDAR HEATMAP
# ═══════════════════════════════════════════════════════════════════════════════
def build_calendar_heatmap(parent, year: int, month: int, attendance_days: set[str], theme, lang: str = "en"):
    """Build a tkinter calendar grid showing present/absent days."""
    import calendar
    cal = calendar.Calendar()
    month_days = cal.monthdayscalendar(year, month)
    month_name = "jan" if month == 1 else "feb" if month == 2 else "mar" if month == 3 else                  "apr" if month == 4 else "may" if month == 5 else "jun" if month == 6 else                  "jul" if month == 7 else "aug" if month == 8 else                  "sep" if month == 9 else "oct" if month == 10 else                  "nov" if month == 11 else "dec"

    hdr = tk.Label(parent, text=f"{month_name} {year}", bg=theme.bg, fg=theme.text,
                   font=("Segoe UI", 12, "bold"))
    hdr.pack(anchor="w", pady=(0, 6))

    grid = tk.Frame(parent, bg=theme.bg)
    grid.pack(fill="x")

    days = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
    for i, d in enumerate(days):
        tk.Label(grid, text=d, bg=theme.bg, fg=theme.muted, font=("Segoe UI", 8), width=4).grid(row=0, column=i, padx=2, pady=2)

    for r, week in enumerate(month_days, 1):
        for c, day in enumerate(week):
            if day == 0:
                tk.Label(grid, text="", bg=theme.bg, width=4).grid(row=r, column=c)
                continue
            day_str = f"{year:04d}-{month:02d}-{day:02d}"
            if day_str in attendance_days:
                bg = theme.calendar_present
                fg = "#ffffff"
            else:
                bg = theme.calendar_neutral
                fg = theme.text
            lbl = tk.Label(grid, text=str(day), bg=bg, fg=fg, font=("Segoe UI", 9), width=4, height=2)
            lbl.grid(row=r, column=c, padx=2, pady=2)

# Type scale used across the rebuilt 2.1 pages (kept as plain tuples so the
# existing tk widgets that hard-code their own font don't need touching).
F_H1     = ("Segoe UI", 22, "bold")
F_H2     = ("Segoe UI", 15, "bold")
F_H3     = ("Segoe UI", 12, "bold")
F_BODY   = ("Segoe UI", 10)
F_BODY_B = ("Segoe UI", 10, "bold")
F_SMALL  = ("Segoe UI", 9)
F_SMALL_B= ("Segoe UI", 9, "bold")
F_TINY   = ("Segoe UI", 8)
F_MONO   = ("Consolas", 9)


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def safe_int(v, fallback=0):
    try: return int(v)
    except (TypeError, ValueError): return fallback

def safe_float(v, fallback=1.0):
    try: return float(v)
    except (TypeError, ValueError): return fallback

def col_num(name: str) -> int:
    return COLUMNS.index(name) + 1

def photo_col_letter() -> str:
    return get_column_letter(PHOTO_COLUMN) if get_column_letter else "K"

def clean_phone(v: str) -> str:
    return v.strip().replace(" ", "")

def parse_uid_num(uid: str) -> int:
    m = re.search(r"(\d+)$", uid or "")
    return int(m.group(1)) if m else 0

def open_camera(index=0):
    if cv2 is None:
        return None
    devnull_fd = os.open(os.devnull, os.O_WRONLY)
    old_fd = os.dup(2)
    os.dup2(devnull_fd, 2)
    os.close(devnull_fd)
    try:
        cap = cv2.VideoCapture(index, cv2.CAP_DSHOW) if os.name == "nt" else cv2.VideoCapture(index)
        if cap is None or not cap.isOpened():
            cap = cv2.VideoCapture(index)
        return cap if cap.isOpened() else None
    except Exception:
        return None
    finally:
        os.dup2(old_fd, 2)
        os.close(old_fd)

def ensure_dirs():
    for d in (DATA_DIR, FACES_DIR, UNKNOWNS_DIR, BACKUPS_DIR, EXPORTS_DIR, LOGS_DIR,
              ASSETS_DIR, WHITEBOARD_DIR, NOTES_DIR):
        d.mkdir(parents=True, exist_ok=True)

def _face_recognition_looks_real(mod) -> bool:
    # Catches the case where a local file/folder named face_recognition.py
    # (or a broken face_recognition_models install) shadows the real package.
    return (
        mod is not None
        and hasattr(mod, "face_locations") and callable(mod.face_locations)
        and hasattr(mod, "face_encodings") and callable(mod.face_encodings)
        and hasattr(mod, "face_distance") and callable(mod.face_distance)
    )

def ensure_face_recognition() -> bool:
    global face_recognition, FACE_RECOGNITION_IMPORT_ERROR
    if face_recognition is not None:
        if _face_recognition_looks_real(face_recognition):
            return True
        mod_file = getattr(face_recognition, "__file__", "?")
        FACE_RECOGNITION_IMPORT_ERROR = (
            f"'face_recognition' resolved to {mod_file!r} but is missing the expected "
            f"functions. This usually means a local file/folder named face_recognition.py "
            f"or face_recognition/ is shadowing the real installed package, or that "
            f"'face_recognition_models' failed to install correctly."
        )
        return False
    try:
        import importlib
        candidate = importlib.import_module("face_recognition")
        if not _face_recognition_looks_real(candidate):
            mod_file = getattr(candidate, "__file__", "?")
            FACE_RECOGNITION_IMPORT_ERROR = (
                f"'face_recognition' resolved to {mod_file!r} but is missing the expected "
                f"functions. This usually means a local file/folder named face_recognition.py "
                f"or face_recognition/ is shadowing the real installed package, or that "
                f"'face_recognition_models' failed to install correctly."
            )
            return False
        face_recognition = candidate
        return True
    except BaseException as e:
        FACE_RECOGNITION_IMPORT_ERROR = str(e) or type(e).__name__
        return False


# ═══════════════════════════════════════════════════════════════════════════════
#  SETTINGS  (persisted to JSON)
# ═══════════════════════════════════════════════════════════════════════════════
class Settings:
    _defaults = {
        "late_time": "08:32",
        "threshold": 0.50,
        "xp_ontime": 10,
        "xp_late": 2,
        "school_name": "",
        "smtp_sender": "",
        "smtp_password": "",
        "auto_notify_email": False,
        "auto_notify_whatsapp": False,
        "auto_message_template": (
            "Dear Parent,\n\n"
            "Your ward {name} (Roll No: {roll}) has arrived at school "
            "and was marked {status} at {time} on {date}.\n\n"
            "Attendance XP earned: +{xp}\n\n"
            "— {school}"
        ),
        "language": "en",
        "theme_dark": True,
        "voice_feedback": True,
        "liveness_check": True,
        "camera_index": 0,
        "api_server": False,
        "absent_notify_enabled": False,
        "absent_notify_time": "09:30",
        "absent_notify_email": True,
        "absent_notify_whatsapp": False,
        "absent_message_template": (
            "Dear Parent,\n\n"
            "This is to inform you that your ward {name} (Roll No: {roll}) "
            "has not been marked present at school today, {date}, as of {time}.\n\n"
            "If this is an error or your child is on approved leave, please "
            "contact the school office.\n\n"
            "— {school}"
        ),
        "_absent_notify_last_run_date": "",
    }

    def __init__(self):
        self._data = dict(self._defaults)
        self.load()

    def load(self):
        if SETTINGS_FILE.exists():
            try:
                with SETTINGS_FILE.open() as f:
                    saved = json.load(f)
                self._data.update(saved)
            except Exception:
                pass
        if self._data.get("language") not in LANG_NAMES:
            self._data["language"] = "en"

    def save(self):
        with SETTINGS_FILE.open("w") as f:
            json.dump(self._data, f, indent=2)

    def __getattr__(self, key):
        if key.startswith("_"):
            raise AttributeError(key)
        return self._data.get(key)

    def set(self, key, value):
        self._data[key] = value

    @property
    def late_time_obj(self) -> dtime:
        try:
            h, m = self._data["late_time"].split(":")
            return dtime(int(h), int(m))
        except Exception:
            return dtime(8, 32)

    @property
    def language(self) -> str:
        return self._data.get("language", "en")

    @property
    def theme_dark(self) -> bool:
        return self._data.get("theme_dark", True)

    @property
    def voice_feedback(self) -> bool:
        return self._data.get("voice_feedback", True)

    @property
    def liveness_check(self) -> bool:
        return self._data.get("liveness_check", True)

    @property
    def camera_index(self) -> int:
        return safe_int(self._data.get("camera_index"), 0)

    @property
    def api_server(self) -> bool:
        return self._data.get("api_server", False)


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL STORE
# ═══════════════════════════════════════════════════════════════════════════════
class AttendanceStore:
    def __init__(self, settings: Settings):
        self.settings = settings
        ensure_dirs()
        self.known_uids: list[str] = []
        self.known_encodings: list = []
        self._init_excel()
        self._init_encodings()
        self._init_audit_log()

    # ── internal helpers ──────────────────────────────────────────────────────
    def _init_excel(self):
        if EXCEL_FILE.exists(): return
        if Workbook is None: return
        wb = Workbook()
        ws = wb.active
        ws.title = "Attendance"
        ws.append(COLUMNS)
        self._style_header(ws)
        self._apply_widths(ws)
        ws.cell(1, PHOTO_COLUMN).value = "Photo"
        if get_column_letter:
            ws.column_dimensions[photo_col_letter()].width = 15
        wb.save(EXCEL_FILE)
        wb.close()

    def _style_header(self, ws):
        if PatternFill is None: return
        hdr_fill = PatternFill("solid", fgColor="1E3A5F")
        hdr_font = Font(bold=True, color="FFFFFF", size=11)
        for col in range(1, PHOTO_COLUMN + 1):
            cell = ws.cell(1, col)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    def _apply_widths(self, ws):
        if get_column_letter is None: return
        for i, name in enumerate(COLUMNS, 1):
            ws.column_dimensions[get_column_letter(i)].width = EXCEL_COLUMN_WIDTHS.get(name, 14)

    def _init_encodings(self):
        if not ENCODINGS_FILE.exists():
            with ENCODINGS_FILE.open("wb") as f:
                pickle.dump({"uids": [], "encodings": []}, f)
        else:
            self.load_encodings()

    def _init_audit_log(self):
        if not AUDIT_LOG.exists():
            AUDIT_LOG.parent.mkdir(parents=True, exist_ok=True)
            with AUDIT_LOG.open("w", newline="") as f:
                csv.writer(f).writerow([
                    "Date", "Time", "UID", "Name", "Roll No", "Grade",
                    "House", "Status", "XP Earned", "Streak"
                ])

    def _safe_load(self):
        import tempfile
        tmp = Path(tempfile.mktemp(suffix=".xlsx"))
        try:
            shutil.copy2(EXCEL_FILE, tmp)
            return load_workbook(tmp, read_only=True, data_only=True), tmp
        except Exception:
            try:
                return load_workbook(EXCEL_FILE, read_only=True, data_only=True), None
            except Exception:
                return None, None

    def _kill_excel(self):
        if os.name != "nt": return
        try:
            import subprocess as _sp
            _sp.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True, timeout=5)
            _time.sleep(0.6)
        except Exception: pass

    # ── public read ───────────────────────────────────────────────────────────
    def load_students(self) -> list[dict]:
        if not EXCEL_FILE.exists():
            self.ensure_file_present_and_valid()
        if load_workbook is None or not EXCEL_FILE.exists(): return []
        with _excel_lock:
            wb, tmp = self._safe_load()
            if wb is None: return []
            ws = wb.active
            students = []
            try:
                rows_iter = ws.iter_rows(min_row=1, values_only=True)
                header_row = next(rows_iter, None)
            except StopIteration:
                header_row = None
            header = self._normalize_header(header_row)
            for row in rows_iter:
                if not row: continue
                rec = dict(zip(header, row))
                uid_val = rec.get("UID")
                if not uid_val: continue
                s = {col: rec.get(col) for col in COLUMNS}
                s["Sl.No"]               = safe_int(s.get("Sl.No"))
                s["No. of Days Present"] = safe_int(s.get("No. of Days Present"))
                s["XP"]                  = safe_int(s.get("XP"))
                s["Multiplier"]          = safe_float(s.get("Multiplier"), 1.0)
                s["Streak"]              = safe_int(s.get("Streak"))
                if not s.get("House"):
                    s["House"] = "None"
                students.append(s)
            try: wb.close()
            except: pass
            if tmp and tmp.exists():
                try: tmp.unlink()
                except: pass
            return students

    @staticmethod
    def _normalize_header(header_row) -> list[str]:
        """
        Maps the on-disk header row to canonical COLUMNS names, so older
        files (saved before 'House'/'Streak' existed) still read correctly
        by NAME instead of by position. This is the fix for the bug where
        a stale-schema file made Phone numbers appear under the House
        column — load_students() previously zipped values positionally
        against the CURRENT COLUMNS list regardless of what was actually
        on disk.
        """
        aliases = {
            "phone": "Parent Phone", "parent_phone": "Parent Phone",
            "email": "Parent Email", "parent_email": "Parent Email",
            "days present": "No. of Days Present", "days": "No. of Days Present",
            "roll_no": "Roll No", "roll": "Roll No",
        }
        if not header_row:
            return list(COLUMNS)
        out = []
        for cell in header_row:
            name = str(cell).strip() if cell else ""
            if name in COLUMNS:
                out.append(name)
            elif name.lower() in aliases:
                out.append(aliases[name.lower()])
            elif name == "Photo":
                out.append("Photo")
            else:
                out.append(name or f"_col{len(out)}")
        return out

    def validate_schema(self) -> dict:
        """
        Checks the on-disk header row against the COLUMNS the app expects.
        Returns {"ok": bool, "missing": [...], "extra": [...], "order_ok": bool}.
        Used on every load to warn the teacher (instead of silently reading
        misaligned data) and to drive auto-migration.
        """
        if load_workbook is None or not EXCEL_FILE.exists():
            return {"ok": False, "missing": list(COLUMNS), "extra": [], "order_ok": False, "reason": "missing_file"}
        try:
            wb = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
            ws = wb.active
            first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            wb.close()
        except Exception as e:
            return {"ok": False, "missing": list(COLUMNS), "extra": [], "order_ok": False, "reason": f"unreadable: {e}"}

        header = [str(c).strip() if c else "" for c in (first_row or [])]
        header_cols = [h for h in header if h and h != "Photo"]
        missing = [c for c in COLUMNS if c not in header_cols]
        extra   = [h for h in header_cols if h not in COLUMNS]
        expected_prefix = COLUMNS[:len(header_cols)] if len(header_cols) <= len(COLUMNS) else COLUMNS
        order_ok = header_cols[:len(expected_prefix)] == expected_prefix
        return {
            "ok": not missing and not extra and order_ok,
            "missing": missing, "extra": extra, "order_ok": order_ok,
            "reason": "header_mismatch" if (missing or extra or not order_ok) else "ok",
        }

    def migrate_schema(self) -> bool:
        """
        Rewrites the Excel file onto the CURRENT COLUMNS layout, preserving
        every value by NAME (not position), backing up the old file first.
        Safe to call even if the file already matches — it's a no-op cost
        but never destructive since load_students() now reads by name too.
        """
        if load_workbook is None or Workbook is None or not EXCEL_FILE.exists():
            return False
        try:
            old_students = self.load_students()
        except Exception:
            return False

        # back up the pre-migration file first
        try:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(EXCEL_FILE, BACKUPS_DIR / f"FaceAttend_premigration_{stamp}.xlsx")
        except Exception:
            pass

        with _excel_lock:
            self._kill_excel()
            wb = Workbook()
            ws = wb.active
            ws.title = "Attendance"
            ws.append(COLUMNS)
            self._style_header(ws)
            self._apply_widths(ws)
            ws.cell(1, PHOTO_COLUMN).value = "Photo"
            if get_column_letter:
                ws.column_dimensions[photo_col_letter()].width = 15
            for i, s in enumerate(old_students, 2):
                ws.append([s.get(c, "") for c in COLUMNS])
                self._stripe_row(ws, i)
                uid = str(s.get("UID", ""))
                fp  = FACES_DIR / f"{uid}.jpg"
                if fp.exists():
                    self._attach_photo(ws, i, fp)
            wb.save(EXCEL_FILE)
            wb.close()
        return True

    def ensure_file_present_and_valid(self) -> dict:
        """
        Call this on startup and before any risky operation. Handles two
        distinct failure modes the teacher reported:
          1) The .xlsx file was deleted outright while the app wasn't
             watching — auto-restore from the newest backup if one exists,
             otherwise create a brand-new empty sheet. Always returns a
             warning message so the UI can toast it (never silent).
          2) The .xlsx file exists but its header row doesn't match the
             current schema (the House/Phone column-shift bug) — auto
             migrate it onto the correct layout, preserving all data by
             column NAME.
        Returns {"action": "none"|"restored"|"created"|"migrated", "message": str|None}
        """
        if not EXCEL_FILE.exists():
            backups = sorted(BACKUPS_DIR.glob("FaceAttend_backup_*.xlsx"),
                              key=lambda p: p.stat().st_mtime, reverse=True) if BACKUPS_DIR.exists() else []
            if backups:
                try:
                    shutil.copy2(backups[0], EXCEL_FILE)
                    return {"action": "restored",
                            "message": f"Excel file was missing — restored from backup ({backups[0].name})."}
                except Exception:
                    pass
            self._init_excel()
            return {"action": "created",
                    "message": "Excel file was missing — created a new empty sheet."}

        check = self.validate_schema()
        if not check["ok"] and check.get("reason") == "header_mismatch":
            migrated = self.migrate_schema()
            if migrated:
                return {"action": "migrated",
                        "message": "Excel column layout was out of date (e.g. an older file without "
                                    "House/Streak columns) — it has been automatically corrected. "
                                    "A backup of the original file was saved first."}
        return {"action": "none", "message": None}

    def students_by_uid(self) -> dict[str, dict]:
        return {str(s["UID"]): s for s in self.load_students()}

    def load_encodings(self):
        self.known_uids = []
        self.known_encodings = []
        if not ENCODINGS_FILE.exists(): return
        try:
            with ENCODINGS_FILE.open("rb") as f:
                data = pickle.load(f)
            self.known_uids      = list(data.get("uids", []))
            self.known_encodings = list(data.get("encodings", []))
        except Exception: pass

    def save_encodings(self):
        with ENCODINGS_FILE.open("wb") as f:
            pickle.dump({"uids": self.known_uids, "encodings": self.known_encodings}, f)

    def next_uid(self) -> str:
        used = [parse_uid_num(str(s["UID"])) for s in self.load_students()]
        used += [parse_uid_num(u) for u in self.known_uids]
        return f"F{max(used, default=0)+1:03d}"

    # ── public write ─────────────────────────────────────────────────────────
    def add_student(self, roll_no, uid, name, grade, house, phone, email, encoding, face_path=None):
        self.ensure_file_present_and_valid()
        with _excel_lock:
            self._kill_excel()
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            serial = max(ws.max_row, 1)
            ws.append([serial, roll_no, uid, name, grade, house or "None", phone, email, 0, 0, 1.0, 0])
            row_idx = ws.max_row
            self._attach_photo(ws, row_idx, face_path)
            self._stripe_row(ws, row_idx)
            wb.save(EXCEL_FILE)
            wb.close()
        self.known_uids.append(uid)
        self.known_encodings.append(encoding)
        self.save_encodings()

    def edit_student(self, uid, roll_no, name, grade, house, phone, email):
        self.ensure_file_present_and_valid()
        with _excel_lock:
            self._kill_excel()
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            found = False
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, col_num("UID")).value) == uid:
                    ws.cell(r, col_num("Roll No")).value        = roll_no
                    ws.cell(r, col_num("Name")).value           = name
                    ws.cell(r, col_num("Grade")).value          = grade
                    ws.cell(r, col_num("House")).value          = house or "None"
                    ws.cell(r, col_num("Parent Phone")).value   = phone
                    ws.cell(r, col_num("Parent Email")).value   = email
                    found = True
                    break
            wb.save(EXCEL_FILE)
            wb.close()
            if not found:
                raise ValueError(f"Student {uid} not found for edit")

    def remove_student(self, uid) -> dict | None:
        self.ensure_file_present_and_valid()
        with _excel_lock:
            self._kill_excel()
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            removed = None
            target = None
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, col_num("UID")).value) == uid:
                    vals = [ws.cell(r, c).value for c in range(1, len(COLUMNS) + 1)]
                    removed = dict(zip(COLUMNS, vals))
                    target = r
                    break
            if target:
                self._remove_images_in_row(ws, target)
                ws.delete_rows(target, 1)
                self._renumber(ws)
                wb.save(EXCEL_FILE)
            wb.close()
        if uid in self.known_uids:
            keep = [i for i, u in enumerate(self.known_uids) if u != uid]
            self.known_uids      = [self.known_uids[i]      for i in keep]
            self.known_encodings = [self.known_encodings[i] for i in keep]
            self.save_encodings()
        fp = FACES_DIR / f"{uid}.jpg"
        if fp.exists(): fp.unlink()
        return removed

    def bulk_remove(self, uids: list[str]) -> int:
        """
        Remove many students in a single workbook open/save cycle (used by
        multi-select Remove and by AI Buddy bulk-remove commands) — far
        faster than calling remove_student() in a loop for a whole class,
        and avoids re-locking/re-opening the Excel file per student.
        """
        uid_set = set(uids)
        if not uid_set:
            return 0
        self.ensure_file_present_and_valid()
        removed_count = 0
        with _excel_lock:
            self._kill_excel()
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            # walk bottom-up so deleting a row doesn't shift the indices
            # of rows still waiting to be checked
            for r in range(ws.max_row, 1, -1):
                if str(ws.cell(r, col_num("UID")).value) in uid_set:
                    self._remove_images_in_row(ws, r)
                    ws.delete_rows(r, 1)
                    removed_count += 1
            if removed_count:
                self._renumber(ws)
                wb.save(EXCEL_FILE)
            wb.close()
        keep = [i for i, u in enumerate(self.known_uids) if u not in uid_set]
        if len(keep) != len(self.known_uids):
            self.known_uids      = [self.known_uids[i]      for i in keep]
            self.known_encodings = [self.known_encodings[i] for i in keep]
            self.save_encodings()
        for uid in uid_set:
            fp = FACES_DIR / f"{uid}.jpg"
            if fp.exists(): fp.unlink()
        return removed_count

    def bulk_set_grades(self, uid_to_grade: dict[str, str]) -> int:
        """
        Set the Grade/Section field for many students at once, in a single
        workbook open/save cycle. Backs both the year-end Promotion tool
        (7C → 8C for the whole school) and ad-hoc multi-select /
        AI-driven section changes.
        """
        if not uid_to_grade:
            return 0
        self.ensure_file_present_and_valid()
        changed = 0
        with _excel_lock:
            self._kill_excel()
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            for r in range(2, ws.max_row + 1):
                uid = str(ws.cell(r, col_num("UID")).value)
                if uid in uid_to_grade:
                    ws.cell(r, col_num("Grade")).value = uid_to_grade[uid]
                    changed += 1
            if changed:
                wb.save(EXCEL_FILE)
            wb.close()
        return changed

    def record_attendance(self, uid: str, is_late: bool, settings: Settings):
        self.ensure_file_present_and_valid()
        with _excel_lock:
            self._kill_excel()
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, col_num("UID")).value) != uid: continue

                days_c   = col_num("No. of Days Present")
                xp_c     = col_num("XP")
                mul_c    = col_num("Multiplier")
                streak_c = col_num("Streak")

                mul    = safe_float(ws.cell(r, mul_c).value, 1.0)
                streak = safe_int(ws.cell(r, streak_c).value) + 1
                base   = safe_int(settings.xp_late, 2) if is_late else safe_int(settings.xp_ontime, 10)
                xp_e   = int(round(base * mul))

                days   = safe_int(ws.cell(r, days_c).value) + 1
                xp_t   = safe_int(ws.cell(r, xp_c).value) + xp_e

                # streak multiplier: every 5 days +0.1, cap at 3.0
                if days % 5 == 0:
                    mul = min(3.0, round(mul + 0.1, 1))

                ws.cell(r, days_c).value   = days
                ws.cell(r, xp_c).value     = xp_t
                ws.cell(r, mul_c).value    = mul
                ws.cell(r, streak_c).value = streak

                vals    = [ws.cell(r, c).value for c in range(1, len(COLUMNS) + 1)]
                student = dict(zip(COLUMNS, vals))
                wb.save(EXCEL_FILE)
                wb.close()

                # audit log
                try:
                    now = datetime.now()
                    with AUDIT_LOG.open("a", newline="") as lf:
                        csv.writer(lf).writerow([
                            now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"),
                            uid, student.get("Name",""), student.get("Roll No",""),
                            student.get("Grade",""), student.get("House",""),
                            "Late" if is_late else "On-time", xp_e, streak
                        ])
                except Exception: pass

                return student, xp_e
            wb.close()
        raise ValueError(f"No student with UID {uid}")

    def export_excel(self, dest: Path) -> Path | None:
        """Create a styled export copy."""
        with _excel_lock:
            wb_src, tmp = self._safe_load()
            if wb_src is None: return None
        if Workbook is None: return None
        wb = Workbook()
        ws = wb.active
        ws.title = "Student Attendance"
        students = self.load_students()
        headers = COLUMNS + ["Photo"]
        ws.append(headers)
        self._style_header(ws)
        self._apply_widths(ws)
        ws.column_dimensions[photo_col_letter()].width = 15

        row_fills = [PatternFill("solid", fgColor="0D1B2E"), PatternFill("solid", fgColor="112035")] if PatternFill else [None, None]
        for i, s in enumerate(students, 2):
            ws.append([s.get(c, "") for c in COLUMNS])
            if row_fills[0]:
                for cc in range(1, len(COLUMNS) + 1):
                    ws.cell(i, cc).fill = row_fills[(i) % 2]
                    ws.cell(i, cc).font = Font(color="E8EDF5")
                    ws.cell(i, cc).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            uid = str(s.get("UID", ""))
            self._attach_photo(ws, i, FACES_DIR / f"{uid}.jpg")

        dest.parent.mkdir(parents=True, exist_ok=True)
        wb.save(dest)
        wb.close()
        try:
            wb_src.close()
        except Exception:
            pass
        if tmp and tmp.exists():
            try: tmp.unlink()
            except: pass
        return dest

    def export_csv(self, dest: Path) -> Path:
        students = self.load_students()
        dest.parent.mkdir(parents=True, exist_ok=True)
        with dest.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=COLUMNS)
            w.writeheader()
            w.writerows(students)
        return dest

    def backup(self) -> Path:
        BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = BACKUPS_DIR / f"FaceAttend_backup_{stamp}.xlsx"
        if EXCEL_FILE.exists():
            shutil.copy2(EXCEL_FILE, dest)
        return dest

    def bulk_import(self, import_path: Path) -> tuple[int, list[str]]:
        """Import students from an xlsx. Expected columns (order flexible):
        Roll No, Name, Grade, House, Parent Phone, Parent Email.
        Returns (added_count, list_of_skipped_reasons)."""
        self.ensure_file_present_and_valid()
        if load_workbook is None:
            return 0, ["openpyxl not installed"]
        wb_in = None
        try:
            wb_in = load_workbook(str(import_path), read_only=True, data_only=True)
            ws_in = wb_in.active
            first_row = next(ws_in.iter_rows(min_row=1, max_row=1))
            header = [str(c.value).strip() if c.value else "" for c in first_row]
            added, skipped = 0, []
            existing_rolls = {str(s.get("Roll No","")).strip() for s in self.load_students()}
            for row in ws_in.iter_rows(min_row=2, values_only=True):
                if not any(row): continue
                data  = dict(zip(header, row))
                roll  = str(data.get("Roll No",      data.get("roll_no",""))).strip()
                name  = str(data.get("Name",         data.get("name",""))).strip()
                grade = str(data.get("Grade",        data.get("grade",""))).strip()
                house = str(data.get("House",        data.get("house","None"))).strip() or "None"
                phone = str(data.get("Parent Phone", data.get("phone",""))).strip()
                email = str(data.get("Parent Email", data.get("email",""))).strip()
                if not roll or not name:
                    skipped.append("Row skipped — missing Roll No or Name"); continue
                if roll in existing_rolls:
                    skipped.append(f"{roll} ({name}) — already exists"); continue
                uid = self.next_uid()
                with _excel_lock:
                    self._kill_excel()
                    wb = load_workbook(EXCEL_FILE)
                    ws = wb.active
                    serial = max(ws.max_row, 1)
                    ws.append([serial, roll, uid, name, grade, house, phone, email, 0, 0, 1.0, 0])
                    self._stripe_row(ws, ws.max_row)
                    wb.save(EXCEL_FILE); wb.close()
                existing_rolls.add(roll)
                added += 1
            return added, skipped
        except Exception as e:
            return 0, [f"Error reading file: {e}"]
        finally:
            if wb_in:
                try: wb_in.close()
                except: pass

    def generate_certificates(self, min_pct: float = 90.0) -> list[dict]:
        """Return list of students eligible for an attendance certificate."""
        school_days_total = max(1, len({
            row[0] for row in self._load_audit_rows() if row
        }))
        results = []
        for s in self.load_students():
            days = safe_int(s.get("No. of Days Present", 0))
            pct  = round(days / school_days_total * 100, 1)
            if pct >= min_pct:
                results.append({**s, "_pct": pct, "_school_days": school_days_total})
        return sorted(results, key=lambda x: -x["_pct"])

    def _load_audit_rows(self) -> list:
        if not AUDIT_LOG.exists(): return []
        rows = []
        try:
            with AUDIT_LOG.open("r", newline="") as f:
                rows = list(csv.reader(f))
        except Exception: pass
        return rows[1:] if rows else []

    def house_stats(self) -> dict[str, dict]:
        """Return {house: {students, present, pct}} for the House leaderboard."""
        all_students = self.load_students()
        scanned = self._load_scanned_today_uids()
        stats: dict[str, dict] = {}
        for house in ("Red", "Blue", "Green", "Orange"):
            members = [s for s in all_students if str(s.get("House","")).strip() == house]
            present = sum(1 for s in members if str(s.get("UID","")) in scanned)
            pct     = round(present / len(members) * 100, 1) if members else 0.0
            stats[house] = {"students": len(members), "present": present, "pct": pct}
        return stats

    def _load_scanned_today_uids(self) -> set[str]:
        if not TODAY_SCANS_FILE.exists(): return set()
        try:
            with open(TODAY_SCANS_FILE, "rb") as f:
                payload = pickle.load(f)
            if payload.get("date") == date.today().isoformat():
                return set(payload.get("uids", []))
        except Exception: pass
        return set()

    # ── static helpers ────────────────────────────────────────────────────────
    @staticmethod
    def _renumber(ws):
        for r in range(2, ws.max_row + 1):
            ws.cell(r, 1).value = r - 1

    @staticmethod
    def _stripe_row(ws, row_idx):
        if PatternFill is None: return
        fill = PatternFill("solid", fgColor="0D1B2E" if row_idx % 2 == 0 else "112035")
        for c in range(1, PHOTO_COLUMN):
            ws.cell(row_idx, c).fill = fill
            ws.cell(row_idx, c).font = Font(color="E8EDF5")

    @staticmethod
    def _attach_photo(ws, row_idx, face_path):
        if ExcelImage is None or face_path is None or not Path(face_path).exists(): return
        try:
            img = ExcelImage(str(face_path))
            img.width  = 72
            img.height = 72
            ws.row_dimensions[row_idx].height = 58
            ws.add_image(img, f"{photo_col_letter()}{row_idx}")
        except Exception: pass

    @staticmethod
    def _remove_images_in_row(ws, row_idx):
        if not hasattr(ws, "_images"):
            return
        try:
            ws._images = [
                im for im in ws._images
                if not (
                    hasattr(im, "anchor") and
                    hasattr(im.anchor, "_from") and
                    getattr(im.anchor._from, "row", None) == row_idx - 1
                )
            ]
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════════════
#  ANIMATED BUTTON  — spring press effect
# ═══════════════════════════════════════════════════════════════════════════════
class AnimBtn(tk.Button):
    """Button with a subtle press-spring animation and colour ripple."""

    def __init__(self, master, anim_color: str | None = None, **kw):
        self._orig_bg  = kw.get("bg", P.blue)
        self._anim_col = anim_color or self._orig_bg
        self._anim_job = None
        super().__init__(master, **kw)
        self.bind("<ButtonPress-1>",   self._on_press,   add="+")
        self.bind("<ButtonRelease-1>", self._on_release, add="+")

    def _on_press(self, _e=None):
        self._cancel()
        try:
            self.configure(relief="sunken", padx=max(0, int(self.cget("padx")) - 1))
            self._ripple(0)
        except tk.TclError:
            pass

    def _on_release(self, _e=None):
        self._cancel()
        try:
            self.configure(relief="flat", padx=int(self.cget("padx")) + 1)
            self._spring_back(0)
        except tk.TclError:
            pass

    def _ripple(self, step):
        if step < 3:
            try:
                self.configure(bg=self._anim_col)
                self._anim_job = self.after(40, lambda: self._ripple(step + 1))
            except tk.TclError:
                pass

    def _spring_back(self, step):
        if step < 4:
            try:
                self._anim_job = self.after(35, lambda: self._spring_back(step + 1))
            except tk.TclError:
                pass
        else:
            try:
                self.configure(bg=self._orig_bg)
            except tk.TclError:
                pass

    def _cancel(self):
        if self._anim_job:
            try: self.after_cancel(self._anim_job)
            except: pass
            self._anim_job = None


# ═══════════════════════════════════════════════════════════════════════════════
#  TOOLTIP  — small delayed label used by the icon nav rail
# ═══════════════════════════════════════════════════════════════════════════════
class Tooltip:
    """Lightweight hover tooltip. Shows after a short delay near the widget."""

    def __init__(self, widget, text: str, delay: int = 450):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._after_id = None
        self._tip = None
        self._enter_id = widget.bind("<Enter>", self._schedule, add="+")
        self._leave_id = widget.bind("<Leave>", self._hide, add="+")
        self._btn_id  = widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _e=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id:
            try: self.widget.after_cancel(self._after_id)
            except: pass
            self._after_id = None

    def _show(self):
        if self._tip is not None:
            return
        try:
            x = self.widget.winfo_rootx() + self.widget.winfo_width() + 8
            y = self.widget.winfo_rooty() + (self.widget.winfo_height() // 2) - 11
        except tk.TclError:
            return
        self._tip = tk.Toplevel(self.widget)
        self._tip.overrideredirect(True)
        self._tip.attributes("-topmost", True)
        self._tip.configure(bg=P.border)
        inner = tk.Frame(self._tip, bg=P.surface2, highlightthickness=1,
                          highlightbackground=P.border)
        inner.pack(padx=1, pady=1)
        tk.Label(inner, text=self.text, bg=P.surface2, fg=P.text,
                 font=F_SMALL_B, padx=10, pady=5).pack()
        self._tip.geometry(f"+{x}+{y}")

    def _hide(self, _e=None):
        self._cancel()
        if self._tip is not None:
            try: self._tip.destroy()
            except tk.TclError: pass
            self._tip = None

    def destroy(self):
        self._hide()
        try:
            self.widget.unbind("<Enter>", self._enter_id)
            self.widget.unbind("<Leave>", self._leave_id)
            self.widget.unbind("<ButtonPress>", self._btn_id)
        except tk.TclError:
            pass


def add_hover(widget: tk.Widget, normal_bg: str, hover_bg: str) -> None:
    """Attach a simple background-swap hover effect to a plain tk widget."""
    widget.bind("<Enter>", lambda _e: widget.configure(bg=hover_bg), add="+")
    widget.bind("<Leave>", lambda _e: widget.configure(bg=normal_bg), add="+")


# ═══════════════════════════════════════════════════════════════════════════════
#  TOAST NOTIFICATION
# ═══════════════════════════════════════════════════════════════════════════════
class Toast:
    """Non-blocking slide-in notification at bottom-right."""
    _queue: list = []
    _showing = False

    @classmethod
    def show(cls, root: tk.Tk, msg: str, kind: str = "info", duration: int = 3000,
             action_label: str | None = None, action_command=None):
        cls._queue.append((root, msg, kind, duration, action_label, action_command))
        if not cls._showing:
            cls._next()

    @classmethod
    def _next(cls):
        if not cls._queue:
            cls._showing = False
            return
        cls._showing = True
        root, msg, kind, duration, action_label, action_command = cls._queue.pop(0)
        cls._display(root, msg, kind, duration, action_label, action_command)

    @classmethod
    def _display(cls, root, msg, kind, duration, action_label=None, action_command=None):
        color = {
            "info":    P.blue,
            "success": P.green,
            "error":   P.red,
            "warning": P.amber,
        }.get(kind, P.blue)
        icon = {"info": "ℹ", "success": "✓", "error": "✕", "warning": "⚠"}.get(kind, "ℹ")

        win = tk.Toplevel(root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(bg=P.surface)
        win.attributes("-alpha", 0.0)

        frm = tk.Frame(win, bg=P.surface, highlightthickness=2, highlightbackground=color)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text=icon, bg=P.surface, fg=color,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=(14, 6), pady=14)
        tk.Label(frm, text=msg, bg=P.surface, fg=P.text,
                 font=("Segoe UI", 10), wraplength=240, justify="left").pack(
                     side="left", padx=(0, 10), pady=14)

        if action_label and action_command:
            def _run_action():
                try: win.destroy()
                except tk.TclError: pass
                try: root.after(60, cls._next)
                except tk.TclError: pass
                action_command()
            btn = tk.Button(frm, text=action_label, command=_run_action,
                            bg=color, fg="white", activebackground=color,
                            activeforeground="white", relief="flat", bd=0,
                            padx=12, pady=6, font=("Segoe UI", 9, "bold"), cursor="hand2")
            btn.pack(side="left", padx=(0, 14), pady=14)

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        sw   = root.winfo_screenwidth()
        sh   = root.winfo_screenheight()
        x    = sw - w - 22
        y_target = sh - h - 52

        win.geometry(f"{w}x{h}+{x}+{sh}")

        def slide(step, start_y, end_y):
            if step > 15:
                try:
                    win.attributes("-alpha", 1.0)
                    root.after(duration, lambda: fade_out(0))
                except tk.TclError:
                    pass
                return
            try:
                frac = step / 15
                ease = 1 - (1 - frac) ** 3
                cy   = int(start_y + (end_y - start_y) * ease)
                win.geometry(f"{w}x{h}+{x}+{cy}")
                win.attributes("-alpha", min(1.0, frac * 2))
                root.after(18, lambda: slide(step + 1, start_y, end_y))
            except tk.TclError:
                pass

        def fade_out(step):
            if step > 10:
                try: win.destroy()
                except tk.TclError: pass
                try: root.after(80, cls._next)
                except tk.TclError: pass
                return
            try:
                win.attributes("-alpha", max(0.0, 1.0 - step / 10))
                root.after(30, lambda: fade_out(step + 1))
            except tk.TclError:
                pass

        slide(0, sh, y_target)


# ═══════════════════════════════════════════════════════════════════════════════
#  STAT CARD
# ═══════════════════════════════════════════════════════════════════════════════
class StatCard(tk.Frame):
    def __init__(self, master, title, value, icon, tint, accent, subtitle=""):
        super().__init__(master, bg=tint, highlightthickness=1, highlightbackground=P.border_soft)

        chip_wrap = tk.Frame(self, bg=tint)
        chip_wrap.pack(anchor="w", padx=16, pady=(14, 0))
        chip = tk.Label(chip_wrap, text=icon, bg=accent, fg="#0b0e14",
                         font=("Segoe UI", 13), width=2, height=1)
        chip.pack()

        tk.Label(self, text=title, bg=tint, fg=P.muted,
                 font=F_SMALL_B).pack(anchor="w", padx=16, pady=(10, 0))
        self.val = tk.Label(self, text=str(value), bg=tint, fg=P.text,
                            font=("Segoe UI", 25, "bold"))
        self.val.pack(anchor="w", padx=16, pady=(2, 0))
        self.sub = tk.Label(self, text=subtitle, bg=tint, fg=P.faint, font=F_TINY)
        self.sub.pack(anchor="w", padx=16, pady=(0, 12))

    def set(self, v):
        try: self.val.configure(text=str(v))
        except tk.TclError: pass

    def set_subtitle(self, text):
        try: self.sub.configure(text=text)
        except tk.TclError: pass


# ═══════════════════════════════════════════════════════════════════════════════
#  ADD/EDIT STUDENT DIALOG
# ═══════════════════════════════════════════════════════════════════════════════
class StudentDialog(tk.Toplevel):
    def __init__(self, app, student: dict | None = None):
        super().__init__(app.root)
        self.app = app
        self.student = student
        self.is_edit  = student is not None
        self.captured_frame    = None
        self.captured_encoding = None
        self.preview_image     = None
        self._mw_binding     = None

        title_text = "Edit Student" if self.is_edit else "Add New Student"
        self.title(title_text)
        self.configure(bg=P.bg)
        self.resizable(False, False)
        self.transient(app.root)
        self.grab_set()

        # cap to 88% of screen height, max 700px so buttons always show
        screen_h = self.winfo_screenheight()
        dlg_h    = min(700, int(screen_h * 0.88))
        self.geometry(f"520x{dlg_h}")
        self.maxsize(520, dlg_h)

        self._build()

        if self.is_edit:
            self.roll_var.set(str(student.get("Roll No", "")))
            self.name_var.set(str(student.get("Name", "")))
            self.grade_var.set(str(student.get("Grade", "")))
            house_val = str(student.get("House", "None")).strip() or "None"
            self.house_var.set(house_val if house_val in HOUSES else "None")
            self.phone_var.set(str(student.get("Parent Phone", "")))
            self.email_var.set(str(student.get("Parent Email", "")))
            uid = str(student.get("UID", ""))
            fp  = FACES_DIR / f"{uid}.jpg"
            if fp.exists() and Image:
                img = Image.open(fp)
                img.thumbnail((490, 200))
                self.preview_image = ImageTk.PhotoImage(img)
                self.preview_lbl.configure(image=self.preview_image, text="")

    def _build(self):
        # ── scrollable shell ──────────────────────────────────────────────────
        outer = tk.Frame(self, bg=P.bg)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg=P.bg, highlightthickness=0)
        vsb    = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        shell = tk.Frame(canvas, bg=P.bg)
        win_id = canvas.create_window((0, 0), window=shell, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win_id, width=e.width))
        shell.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        self._mw_binding = "<MouseWheel>"
        canvas.bind_all(self._mw_binding, lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        self.bind("<Destroy>", lambda e: canvas.unbind_all(self._mw_binding))

        pad = tk.Frame(shell, bg=P.bg)
        pad.pack(fill="both", expand=True, padx=22, pady=18)

        # title
        title_text = "✏  Edit Student" if self.is_edit else "➕  Add Student"
        tk.Label(pad, text=title_text, bg=P.bg, fg=P.text,
                 font=F_H2).pack(anchor="w")
        tk.Label(pad,
                 text="Fill in the details below." + ("" if self.is_edit else " Then capture a face."),
                 bg=P.bg, fg=P.muted, font=F_SMALL).pack(anchor="w", pady=(2, 10))

        # form card
        form = tk.Frame(pad, bg=P.surface, highlightthickness=1, highlightbackground=P.border)
        form.pack(fill="x")

        self.roll_var  = tk.StringVar()
        self.name_var  = tk.StringVar()
        self.grade_var = tk.StringVar()
        self.house_var = tk.StringVar(value="None")
        self.phone_var = tk.StringVar()
        self.email_var = tk.StringVar()

        self._field(form, "Roll No",      self.roll_var,  "e.g. 7C04")
        self._field(form, "Name",         self.name_var,  "e.g. Arjun Kumar")
        self._field(form, "Grade",        self.grade_var, "e.g. 7C")

        # house row
        hr = tk.Frame(form, bg=P.surface)
        hr.pack(fill="x", padx=14, pady=(10, 0))
        tk.Label(hr, text="House", bg=P.surface, fg=P.muted,
                 font=F_SMALL_B, width=14, anchor="w").pack(side="left")
        house_cb = ttk.Combobox(hr, textvariable=self.house_var,
                                values=HOUSES, state="readonly",
                                font=F_BODY, width=12)
        house_cb.pack(side="left")
        self.house_dot = tk.Label(hr, text=" ●", bg=P.surface,
                                   fg=P.faint, font=F_H3)
        self.house_dot.pack(side="left", padx=4)
        house_cb.bind("<<ComboboxSelected>>", self._update_house_dot)

        self._field(form, "Parent Phone", self.phone_var, "+919876543210")
        self._field(form, "Parent Email", self.email_var, "parent@example.com")
        tk.Frame(form, bg=P.border, height=1).pack(fill="x", padx=14, pady=(10, 0))

        # compact photo preview
        prev_box = tk.Frame(pad, bg=P.sidebar2, height=120,
                            highlightthickness=1, highlightbackground=P.border)
        prev_box.pack(fill="x", pady=10)
        prev_box.pack_propagate(False)
        self.preview_lbl = tk.Label(prev_box,
                                    text="📷  No face captured yet",
                                    bg=P.sidebar2, fg=P.muted, font=F_BODY)
        self.preview_lbl.pack(fill="both", expand=True)

        self.msg_lbl = tk.Label(pad, text="", bg=P.bg, fg=P.muted,
                                font=F_SMALL, anchor="w")
        self.msg_lbl.pack(fill="x")

        # buttons — always at bottom of scroll content
        btn_row = tk.Frame(pad, bg=P.bg)
        btn_row.pack(fill="x", pady=(8, 4))

        if not self.is_edit:
            AnimBtn(btn_row, text="📷  Capture",
                    command=self.capture_face,
                    bg=P.blue, fg="white", activebackground=P.blue_dark,
                    activeforeground="white", relief="flat", padx=14, pady=10,
                    font=F_BODY_B, cursor="hand2",
                    anim_color=P.blue_glow).pack(side="left")
        else:
            AnimBtn(btn_row, text="📱  QR Code",
                    command=self._show_qr,
                    bg=P.purple, fg="white", activebackground=P.purple,
                    activeforeground="white", relief="flat", padx=14, pady=10,
                    font=F_BODY_B, cursor="hand2",
                    anim_color="#7c3aed").pack(side="left")

        save_lbl = "💾  Save Changes" if self.is_edit else "✅  Save"
        AnimBtn(btn_row, text=save_lbl, command=self.save,
                bg=P.green, fg="white", activebackground=P.green_dark,
                activeforeground="white", relief="flat", padx=14, pady=10,
                font=F_BODY_B, cursor="hand2",
                anim_color="#15803d").pack(side="right")
        AnimBtn(btn_row, text="✕ Cancel", command=self.destroy,
                bg=P.surface2, fg=P.text, activebackground=P.border,
                activeforeground=P.text, relief="flat", padx=12, pady=10,
                font=F_BODY_B, cursor="hand2").pack(side="right", padx=(0, 8))

        canvas.after(50, lambda: canvas.yview_moveto(0.0))

    def _update_house_dot(self, _e=None):
        h   = self.house_var.get()
        clr = HOUSE_COLORS.get(h, P.faint)
        self.house_dot.configure(fg=clr if clr else P.faint)

    def _field(self, master, label, var, hint):
        row = tk.Frame(master, bg=P.surface)
        row.pack(fill="x", padx=14, pady=(8, 0))
        tk.Label(row, text=label, bg=P.surface, fg=P.muted,
                 font=F_SMALL_B, width=14, anchor="w").pack(side="left")
        e = ttk.Entry(row, textvariable=var, font=F_BODY)
        e.pack(side="left", fill="x", expand=True)
        tk.Label(master, text=hint, bg=P.surface, fg=P.faint,
                 font=F_TINY, anchor="w").pack(fill="x", padx=(156, 14), pady=(1, 0))

    def msg(self, text, color=P.muted):
        self.msg_lbl.configure(text=text, fg=color)

    def capture_face(self):
        if not ensure_face_recognition():
            self.msg("face_recognition not available.", P.red); return
        frame = self.app.get_capture_frame()
        if frame is None:
            self.msg("Could not read from camera.", P.red); return
        rgb  = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        locs = face_recognition.face_locations(rgb, model="hog")
        if len(locs) != 1:
            self.msg("Keep exactly one face in frame.", P.red); return
        encs = face_recognition.face_encodings(rgb, locs)
        if not encs:
            self.msg("Face detected but encoding failed.", P.red); return
        t, r, b, l = locs[0]; pad = 45; h, w = frame.shape[:2]
        crop = frame[max(0,t-pad):min(h,b+pad), max(0,l-pad):min(w,r+pad)]
        self.captured_frame    = crop.copy()
        self.captured_encoding = encs[0]
        rgb2  = cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)
        img   = Image.fromarray(rgb2); img.thumbnail((490, 155))
        self.preview_image = ImageTk.PhotoImage(img)
        self.preview_lbl.configure(image=self.preview_image, text="")
        self.msg("Face captured successfully.", P.green)

    def _show_qr(self):
        """Display the student's QR code in a popup."""
        if not self.is_edit:
            return
        uid = str(self.student.get("UID", ""))
        name = str(self.student.get("Name", ""))
        roll = str(self.student.get("Roll No", ""))
        path = generate_student_qr(uid, name, roll)
        if path is None or not path.exists():
            self.msg("QR generation unavailable (install qrcode).", P.red)
            return
        dlg = tk.Toplevel(self)
        dlg.title(f"QR Code — {name}")
        dlg.configure(bg=P.bg)
        dlg.transient(self)
        tk.Label(dlg, text=f"QR Code for {name} ({roll})", bg=P.bg, fg=P.text,
                 font=F_H3).pack(pady=(16, 8))
        if Image:
            img = Image.open(path)
            img.thumbnail((280, 280))
            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(dlg, image=photo, bg=P.bg)
            lbl.image = photo
            lbl.pack(pady=8)
        tk.Label(dlg, text="Scan this code to mark attendance", bg=P.bg, fg=P.muted,
                 font=F_SMALL).pack(pady=(0, 16))
        AnimBtn(dlg, text="✕  Close", command=dlg.destroy,
            bg=P.surface2, fg=P.text, activebackground=P.border,
            activeforeground=P.text, relief="flat", padx=16, pady=8,
            font=F_BODY_B, cursor="hand2").pack(pady=(0, 12))

    def save(self):
        roll  = self.roll_var.get().strip().upper()
        name  = self.name_var.get().strip()
        grade = self.grade_var.get().strip().upper()
        house = self.house_var.get().strip() or "None"
        phone = clean_phone(self.phone_var.get())
        email = self.email_var.get().strip()

        if not all([roll, name, grade, phone, email]):
            play_sound(SOUND_ERROR)
            self.msg("All fields are required.", P.red); return
        if "@" not in email or "." not in email:
            play_sound(SOUND_ERROR)
            self.msg("Invalid parent email.", P.red); return

        if self.is_edit:
            uid = str(self.student.get("UID", ""))
            try:
                self.app.store.edit_student(uid, roll, name, grade, house, phone, email)
            except ValueError as e:
                play_sound(SOUND_ERROR)
                self.msg(str(e), P.red)
                return
            # Generate QR for existing student
            generate_student_qr(uid, name, roll)
            self.app.refresh_all()
            self.app.log(f"Edited student {name} ({uid}).")
            Toast.show(self.app.root, f"{name} updated.", "success")
            self.destroy()
        else:
            if self.captured_encoding is None:
                play_sound(SOUND_ERROR)
                self.msg("Capture a face first.", P.red); return
            uid  = self.app.store.next_uid()
            fp   = FACES_DIR / f"{uid}.jpg"
            cv2.imwrite(str(fp), self.captured_frame)
            self.app.store.add_student(roll, uid, name, grade, house, phone, email, self.captured_encoding, fp)
            # Generate QR for new student
            generate_student_qr(uid, name, roll)
            self.app.refresh_all()
            self.app.log(f"Added {name} ({uid}).")
            play_sound(SOUND_SUCCESS)
            Toast.show(self.app.root, f"{name} added successfully!", "success")
            self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  YEAR-END PROMOTION  (3.0)
# ═══════════════════════════════════════════════════════════════════════════════
def _split_grade(grade: str) -> tuple[int | None, str]:
    """
    Split a Grade/Section string like '7C' into (7, 'C'). Handles two-digit
    grades ('10A' -> (10, 'A')) and grades with no trailing letters
    ('9' -> (9, '')). Returns (None, original) if no leading number is found.
    """
    m = re.match(r"^\s*(\d{1,2})\s*([A-Za-z]*)\s*$", grade or "")
    if not m:
        return None, (grade or "").strip()
    return int(m.group(1)), m.group(2).strip().upper()


def suggest_next_grade(grade: str) -> str:
    """
    The default promotion suggestion: bump the numeric grade by one and
    keep the section letter — '7C' -> '8C'. Grades FaceAttend can't parse
    (no leading number) are returned unchanged so the teacher can fill
    them in manually.
    """
    num, letters = _split_grade(grade)
    if num is None:
        return grade
    return f"{num + 1}{letters}"


class PromotionDialog(tk.Toplevel):
    """
    Year-end promotion tool. Every distinct Grade/Section currently in the
    roster (e.g. '6A', '6B', '7C' ...) gets one row here, pre-filled with
    the auto-suggested next grade ('7C' -> '8C'). The teacher can edit any
    individual row before applying — e.g. if a section is being merged,
    split, or renamed rather than a simple +1 bump — and one click applies
    every mapping to the whole school in a single bulk write, instead of
    the teacher hand-editing each student's Grade field one at a time.
    Leaving a row blank marks that group for removal (e.g. an outgoing
    graduating class) instead of a grade change, after a confirmation
    that names exactly who will be removed.
    """

    def __init__(self, app: "FaceAttendApp"):
        self.app = app
        super().__init__(app.root)
        self.title("Promote Students — New Academic Year")
        self.configure(bg=P.bg)
        self.geometry("620x560")
        self.transient(app.root)
        self.grab_set()

        students = app.store.load_students()
        self._by_grade: dict[str, list[dict]] = {}
        for s in students:
            g = str(s.get("Grade", "")).strip().upper()
            self._by_grade.setdefault(g, []).append(s)
        self._groups = sorted(self._by_grade.keys(), key=lambda g: _split_grade(g))

        tk.Label(self, text="🎓  Promote Students", bg=P.bg, fg=P.text,
                 font=F_H2).pack(anchor="w", padx=20, pady=(18, 2))
        tk.Label(self, text="Review the suggested next grade for each section below, "
                             "edit anything that isn't a simple +1 move, then Apply.",
                 bg=P.bg, fg=P.muted, font=F_SMALL, wraplength=560, justify="left"
                 ).pack(anchor="w", padx=20, pady=(0, 12))

        if not self._groups:
            tk.Label(self, text="No students enrolled yet — nothing to promote.",
                     bg=P.bg, fg=P.muted, font=F_BODY).pack(padx=20, pady=30)
            AnimBtn(self, text="Close", command=self.destroy,
                bg=P.surface2, fg=P.text, activebackground=P.border,
                activeforeground=P.text, relief="flat", padx=16, pady=8,
                font=F_BODY_B, cursor="hand2").pack(pady=12)
            return

        # header row
        hdr = tk.Frame(self, bg=P.sidebar2)
        hdr.pack(fill="x", padx=20)
        tk.Label(hdr, text="Current", bg=P.sidebar2, fg=P.muted, font=F_SMALL_B,
                  width=12, anchor="w").pack(side="left", padx=(8, 0), pady=6)
        tk.Label(hdr, text="Students", bg=P.sidebar2, fg=P.muted, font=F_SMALL_B,
                  width=10, anchor="w").pack(side="left", pady=6)
        tk.Label(hdr, text="→  New Grade/Section (blank = remove)", bg=P.sidebar2,
                  fg=P.muted, font=F_SMALL_B, anchor="w").pack(side="left", pady=6)

        # scrollable body
        body_wrap = tk.Frame(self, bg=P.bg)
        body_wrap.pack(fill="both", expand=True, padx=20, pady=(4, 8))
        canvas = tk.Canvas(body_wrap, bg=P.bg, highlightthickness=0)
        vsb = ttk.Scrollbar(body_wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=P.bg)
        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._target_vars: dict[str, tk.StringVar] = {}
        for i, g in enumerate(self._groups):
            row = tk.Frame(inner, bg=P.surface if i % 2 == 0 else P.surface2)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=g or "(blank)", bg=row["bg"], fg=P.text,
                      font=F_BODY_B, width=12, anchor="w").pack(side="left", padx=(8, 0), pady=8)
            tk.Label(row, text=str(len(self._by_grade[g])), bg=row["bg"], fg=P.muted,
                      font=F_BODY, width=10, anchor="w").pack(side="left", pady=8)
            var = tk.StringVar(value=suggest_next_grade(g))
            self._target_vars[g] = var
            ttk.Entry(row, textvariable=var, font=F_BODY, width=14).pack(
                side="left", padx=(0, 8), pady=6)

        # footer
        foot = tk.Frame(self, bg=P.bg)
        foot.pack(fill="x", padx=20, pady=(0, 18))
        AnimBtn(foot, text="✕  Cancel", command=self.destroy,
            bg=P.surface2, fg=P.text, activebackground=P.border,
            activeforeground=P.text, relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2").pack(side="left")
        AnimBtn(foot, text="✓  Apply Promotion", command=self._apply,
            bg=P.green, fg="white", activebackground=P.green_dark,
            activeforeground="white", relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color=P.green_dark).pack(side="right")

    def _apply(self):
        grade_changes: dict[str, str] = {}   # uid -> new grade
        remove_names: list[str] = []
        remove_uids: list[str] = []
        moved_groups = 0
        for g in self._groups:
            target = self._target_vars[g].get().strip().upper()
            group_students = self._by_grade[g]
            if target == g:
                continue   # no change requested for this group
            if not target:
                for s in group_students:
                    remove_uids.append(str(s.get("UID", "")))
                    remove_names.append(str(s.get("Name", "")))
                continue
            moved_groups += 1
            for s in group_students:
                grade_changes[str(s.get("UID", ""))] = target

        if not grade_changes and not remove_uids:
            Toast.show(self.app.root, "No changes to apply — every section was left as-is.", "info")
            return

        summary_lines = []
        if grade_changes:
            summary_lines.append(f"• {len(grade_changes)} student(s) across {moved_groups} "
                                  f"section(s) will be moved to their new grade.")
        if remove_uids:
            preview = ", ".join(remove_names[:6])
            more = "" if len(remove_names) <= 6 else f", +{len(remove_names) - 6} more"
            summary_lines.append(f"• {len(remove_uids)} student(s) will be REMOVED "
                                  f"(blank target): {preview}{more}")
        if not messagebox.askyesno("Confirm Promotion",
                                    "Apply the following?\n\n" + "\n".join(summary_lines)):
            return

        if grade_changes:
            self.app.store.bulk_set_grades(grade_changes)
        if remove_uids:
            self.app.store.bulk_remove(remove_uids)
            for uid in remove_uids:
                self.app.scanned_today.discard(uid)
            self.app._save_today_scans()

        self.app.students_by_uid = self.app.store.students_by_uid()
        self.app.refresh_all()
        play_sound(SOUND_SUCCESS)
        total = len(grade_changes) + len(remove_uids)
        Toast.show(self.app.root, f"Promotion applied — {total} student(s) updated.", "success", 5000)
        self.app.log(f"Year-end promotion applied: {len(grade_changes)} promoted, "
                      f"{len(remove_uids)} removed.")
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════
class FaceAttendApp:
    def __init__(self, root: tk.Tk):
        self.root    = root
        self.settings = Settings()
        self.store    = AttendanceStore(self.settings)
        self.students_by_uid: dict[str, dict] = {}
        self.scanned_today: set[str] = set()
        self.camera       = None
        self.scanning     = False
        self.frame_count  = 0
        self._sort_col       = None
        self._sort_ascending = True
        self._unknown_last_saved: float = 0.0
        self.last_frame   = None
        self.last_results: list[dict] = []
        self.camera_photo = None
        self.table_photos: list = []
        self.search_var   = tk.StringVar()
        self._search_after= None
        self.buddy_busy   = False
        self.last_dup_log: dict[str, datetime] = {}
        # ── 2.4 new systems ───────────────────────────────────────────────
        self.schedule     = ScheduleManager()
        self.badges       = BadgeSystem(self.store)
        self.encryption   = EncryptionManager()
        self.liveness     = LivenessDetector()
        self.api_server   = APIServer(self)
        self._voice_enabled = self.settings.voice_feedback
        self._liveness_enabled = self.settings.liveness_check
        self._qr_mode     = False
        self._cameras     = list_cameras()
        self._theme_dark  = self.settings.theme_dark
        self._lang        = self.settings.language if self.settings.language in LANG_NAMES else "en"
        set_theme(self._theme_dark, self._lang)
        # start API if enabled
        if self.settings.api_server:
            self.api_server.start()

        root.title(f"FaceAttend {APP_VERSION}  ·  Smart Attendance System")
        root.geometry("1440x860")
        root.minsize(1100, 720)
        if os.name == "nt":
            root.state("zoomed")
        root.configure(bg=P.bg)
        if ICON_FILE.exists():
            try:
                root.iconbitmap(default=str(ICON_FILE))
            except Exception:
                pass

        self._style()
        self._build()
        self.scanned_today = self._load_today_scans()
        self.refresh_all()

        resilience = self.store.ensure_file_present_and_valid()
        if resilience.get("message"):
            kind = "warning" if resilience["action"] in ("created", "restored") else "info"
            play_sound(SOUND_ERROR if kind == "warning" else SOUND_SUCCESS)
            Toast.show(root, resilience["message"], kind, 6500)
            self.log(resilience["message"])
            self.refresh_all()

        if MISSING_PACKAGES:
            Toast.show(root, f"Missing: {', '.join(set(MISSING_PACKAGES))}", "warning", 5000)
        else:
            self.log(f"FaceAttend {APP_VERSION} ready.")

        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── ttk style ─────────────────────────────────────────────────────────────
    def _style(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("Treeview",
            background=P.surface, foreground=P.text,
            fieldbackground=P.surface, rowheight=70,
            borderwidth=0, font=("Segoe UI", 10))
        s.configure("Treeview.Heading",
            background=P.surface2, foreground=P.muted,
            relief="flat", font=F_SMALL_B, padding=(8, 10))
        s.map("Treeview",
            background=[("selected", P.blue_soft)],
            foreground=[("selected", "white")])
        s.map("Treeview.Heading", background=[("active", P.hover)])
        s.configure("TEntry",   padding=8,  relief="flat",
            fieldbackground=P.surface2, foreground=P.text,
            insertcolor=P.text, bordercolor=P.border,
            lightcolor=P.surface2, darkcolor=P.surface2)
        s.map("TEntry",
            fieldbackground=[("focus", P.hover)],
            bordercolor=[("focus", P.blue)])
        s.configure("TCombobox",
            fieldbackground=P.surface2, foreground=P.text, background=P.surface2,
            arrowcolor=P.muted, bordercolor=P.border,
            lightcolor=P.surface2, darkcolor=P.surface2,
            selectbackground=P.blue, selectforeground="white")
        s.map("TCombobox", fieldbackground=[("readonly", P.surface2)])
        self.root.option_add("*TCombobox*Listbox.background", P.surface2)
        self.root.option_add("*TCombobox*Listbox.foreground", P.text)
        self.root.option_add("*TCombobox*Listbox.selectBackground", P.blue)
        self.root.option_add("*TCombobox*Listbox.font", F_BODY)
        s.configure("Vertical.TScrollbar",
            background=P.surface2, troughcolor=P.bg,
            bordercolor=P.bg, arrowcolor=P.muted, relief="flat")
        s.configure("Horizontal.TScrollbar",
            background=P.surface2, troughcolor=P.bg,
            bordercolor=P.bg, arrowcolor=P.muted, relief="flat")
        s.configure("Splash.Horizontal.TProgressbar",
            troughcolor=P.surface2, background=P.blue,
            bordercolor=P.surface2, lightcolor=P.blue, darkcolor=P.blue_dark)
        s.configure("Accent.Horizontal.TProgressbar",
            troughcolor=P.surface2, background=P.blue,
            bordercolor=P.surface2, lightcolor=P.blue, darkcolor=P.blue, thickness=6)

    # ── layout ────────────────────────────────────────────────────────────────
    NAV_ITEMS = [
        ("\u2302",        "Home",       "home_scanner"),
        ("\U0001F4CA",    "Analytics",  "analytics"),
        ("\U0001F916",    "Bot",        "ai_buddy"),
        ("\u2709",        "Mail",       "messages"),
        ("\U0001F3E0",    "Houses",     "houses"),
        ("\U0001F5CC",    "Whiteboard", "whiteboard"),
        ("\U0001F4DD",    "Notepad",    "notepad"),
        ("\U0001F4CB",    "Console",    "console"),
        ("\u2699",        "Settings",   "settings"),
    ]

    # Tabs that need real screen space (drawing / long-form notes) instead
    # of the narrow 300px page sidebar — for these, the main dashboard
    # (student table etc.) is swapped out for a full-size tool panel.
    FULL_WIDTH_TABS = {"Whiteboard", "Notepad"}

    def _build(self):
        self.root.grid_columnconfigure(0, minsize=84, weight=0)
        self.root.grid_columnconfigure(1, minsize=300, weight=0)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # ── column 0: icon navigation rail ──────────────────────────────────
        self.rail = tk.Frame(self.root, bg=P.rail, width=84)
        self.rail.grid(row=0, column=0, sticky="nsew")
        self.rail.grid_propagate(False)

        glyph = tk.Label(self.rail, text=P.glyph, bg=P.rail, fg=P.blue,
                          font=("Segoe UI", 22))
        glyph.pack(pady=(20, 24))
        Tooltip(glyph, f"FaceAttend {APP_VERSION}  ·  {P.style_name}")

        self._nav_btns = {}
        self._nav_indicators = {}
        nav_wrap = tk.Frame(self.rail, bg=P.rail)
        nav_wrap.pack(fill="x")

        for icon, name, tip in self.NAV_ITEMS:
            row = tk.Frame(nav_wrap, bg=P.rail, height=64)
            row.pack(fill="x")
            row.pack_propagate(False)

            indicator = tk.Frame(row, bg=P.rail, width=4)
            indicator.pack(side="left", fill="y")
            self._nav_indicators[name] = indicator

            btn = tk.Button(row, text=icon, command=lambda n=name: self._switch(n),
                             bg=P.rail, fg=P.muted, activebackground=P.hover,
                             activeforeground=P.text, relief="flat", bd=0,
                             font=("Segoe UI", 19), cursor="hand2")
            btn.pack(side="left", fill="both", expand=True)
            add_hover(btn, P.rail, P.hover)
            Tooltip(btn, _(tip, self._lang))
            self._nav_btns[name] = btn

        # status pill pinned to bottom of the rail (scan state at a glance)
        status_wrap = tk.Frame(self.rail, bg=P.rail)
        status_wrap.pack(side="bottom", pady=18)
        self.rail_status_dot = tk.Label(status_wrap, text="\u25CF", bg=P.rail,
                                         fg=P.faint, font=("Segoe UI", 12))
        self.rail_status_dot.pack()
        self._rail_status_tip = Tooltip(self.rail_status_dot, "Scanner idle")

        # ── column 1: page sidebar ────────────────────────────────────────────
        self.left = tk.Frame(self.root, bg=P.sidebar, width=300)
        self.left.grid(row=0, column=1, sticky="nsew")
        self.left.grid_propagate(False)
        self.left.grid_rowconfigure(0, weight=1)
        self.left.grid_columnconfigure(0, weight=1)

        self._tab_frames = {}
        for _icon, name, _tip in self.NAV_ITEMS:
            f = tk.Frame(self.left, bg=P.sidebar)
            f.grid(row=0, column=0, sticky="nsew")
            self._tab_frames[name] = f

        # ── column 2: main canvas ───────────────────────────────────────────
        # Holds the persistent student dashboard (default) OR, for the
        # Whiteboard/Notepad tabs, a full-size tool panel instead — those
        # need real room to draw/write, not the 300px sidebar column.
        self.right = tk.Frame(self.root, bg=P.bg)
        self.right.grid(row=0, column=2, sticky="nsew")
        self.right.grid_rowconfigure(0, weight=1)
        self.right.grid_columnconfigure(0, weight=1)

        self.dashboard_frame = tk.Frame(self.right, bg=P.bg)
        self.dashboard_frame.grid(row=0, column=0, sticky="nsew")
        self._right_tool_frames: dict[str, tk.Frame] = {}
        for tool_name in self.FULL_WIDTH_TABS:
            f = tk.Frame(self.right, bg=P.bg)
            f.grid(row=0, column=0, sticky="nsew")
            f.grid_remove()
            self._right_tool_frames[tool_name] = f

        def switch(name):
            for n, f in self._tab_frames.items():
                f.grid_remove()
            self._tab_frames[name].grid()
            for n, b in self._nav_btns.items():
                active = n == name
                b.configure(bg=P.hover if active else P.rail,
                            fg=P.text if active else P.muted)
                self._nav_indicators[n].configure(bg=P.blue if active else P.rail)
            # swap the main canvas between the dashboard and a full-width tool
            if name in self.FULL_WIDTH_TABS:
                self.dashboard_frame.grid_remove()
                for n, f in self._right_tool_frames.items():
                    (f.grid() if n == name else f.grid_remove())
            else:
                for f in self._right_tool_frames.values():
                    f.grid_remove()
                self.dashboard_frame.grid()
        self._switch = switch

        self._build_home_tab()
        self._build_right_panel()
        self._build_analytics_tab()
        self._build_bot_tab()
        self._build_mail_tab()
        self._build_houses_tab()
        self._build_whiteboard_tab()
        self._build_notepad_tab()
        self._build_console_tab()
        self._build_settings_tab()
        switch("Home")
        self._tick_clock()
        self._bind_shortcuts()

    def _tick_clock(self):
        """Keep the sidebar masthead clock and dashboard date label fresh."""
        if getattr(self, "_restarting", False):
            return   # old instance being replaced — stop its timer chain
        now = datetime.now()
        if hasattr(self, "clock_lbl"):
            self.clock_lbl.configure(text=now.strftime("%I:%M %p"))
        if hasattr(self, "date_lbl"):
            self.date_lbl.configure(text=now.strftime("%A, %d %B %Y"))
        try:
            self._check_absent_notify_schedule()
        except Exception as e:
            logger.debug("Absent-notify schedule check error: %s", e)
        self.root.after(15000, self._tick_clock)

    def _bind_shortcuts(self):
        """Register global keyboard shortcuts."""
        self.root.bind("<Control-s>", lambda e: (self.toggle_scan(), "break")[-1])
        self.root.bind("<Control-a>", lambda e: (self.open_add_student(), "break")[-1])
        self.root.bind("<Control-f>", lambda e: (self._focus_search(), "break")[-1])
        # Ctrl+1..9 switch tabs
        tab_map = {"1": "Home", "2": "Analytics", "3": "Bot", "4": "Mail",
                   "5": "Houses", "6": "Whiteboard", "7": "Notepad",
                   "8": "Console", "9": "Settings"}
        for num, name in tab_map.items():
            self.root.bind(f"<Control-Key-{num}>", lambda e, n=name: (self._switch(n), "break")[-1])

    def _focus_search(self):
        if hasattr(self, "search_var"):
            # Find the search entry and focus it
            for w in self.root.winfo_children():
                self._focus_search_in(w)

    def _focus_search_in(self, widget):
        if isinstance(widget, ttk.Entry) and str(widget.cget("textvariable")) == str(self.search_var):
            widget.focus_set()
            return True
        for child in widget.winfo_children():
            if self._focus_search_in(child):
                return True
        return False

    def _set_scanning_indicator(self, active: bool) -> None:
        """Update the small status dot on the nav rail."""
        if not hasattr(self, "rail_status_dot"):
            return
        self.rail_status_dot.configure(fg=P.green if active else P.faint)
        if hasattr(self, "_rail_status_tip"):
            self._rail_status_tip.text = "Scanner running" if active else "Scanner idle"


    # ── left: Home tab ────────────────────────────────────────────────────────
    def _build_home_tab(self):
        parent = self._tab_frames["Home"]

        # ── masthead ──────────────────────────────────────────────────────
        hdr = tk.Frame(parent, bg=P.sidebar)
        hdr.pack(fill="x", padx=22, pady=(20, 12))

        title_row = tk.Frame(hdr, bg=P.sidebar)
        title_row.pack(fill="x")
        tk.Label(title_row, text="FaceAttend", bg=P.sidebar, fg=P.text,
                 font=F_H1).pack(side="left")
        chip = tk.Label(title_row, text=f" {APP_VERSION} ", bg=P.blue_soft, fg=P.blue,
                         font=F_SMALL_B)
        chip.pack(side="left", padx=(8, 0), pady=(4, 0))
        self.clock_lbl = tk.Label(title_row, text="", bg=P.sidebar, fg=P.muted,
                                   font=F_SMALL)
        self.clock_lbl.pack(side="right")

        tk.Label(hdr, text="Smart attendance, scanning & rewards",
                 bg=P.sidebar, fg=P.muted, font=F_SMALL).pack(anchor="w", pady=(3, 0))

        # ── camera ────────────────────────────────────────────────────────
        cam_outer = tk.Frame(parent, bg=P.sidebar2, highlightthickness=1,
                              highlightbackground=P.border)
        cam_outer.pack(padx=22, pady=(6, 12), fill="x")
        cam_wrap = tk.Frame(cam_outer, bg="#070d16", width=420, height=300)
        cam_wrap.pack(padx=2, pady=2, fill="x")
        cam_wrap.pack_propagate(False)
        self.cam_lbl = tk.Label(cam_wrap, text="\U0001F4F7  Camera is off",
                                bg="#070d16", fg=P.muted, font=F_H3)
        self.cam_lbl.pack(fill="both", expand=True)

        ctrl = tk.Frame(parent, bg=P.sidebar)
        ctrl.pack(fill="x", padx=22)
        self.scan_btn = AnimBtn(ctrl, text=_("start_scan", self._lang),
            command=self.toggle_scan,
            bg=P.blue, fg="white", activebackground=P.blue_dark,
            activeforeground="white", relief="flat", padx=18, pady=12,
            font=F_BODY_B, cursor="hand2",
            anim_color=P.blue_glow)
        self.scan_btn.pack(side="left", fill="x", expand=True)
        refresh_btn = AnimBtn(ctrl, text="↺  Refresh", command=self.refresh_all,
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=16, pady=12,
            font=F_BODY_B, cursor="hand2")
        refresh_btn.pack(side="left", padx=(10, 0))

        # ── last scan card (will animate on recognition) ────────────────────
        self.result_card = tk.Frame(parent, bg=P.sidebar2,
                                    highlightthickness=1, highlightbackground=P.border)
        self.result_card.pack(fill="x", padx=22, pady=16)
        tk.Label(self.result_card, text="LAST RECOGNITION", bg=P.sidebar2,
                 fg=P.faint, font=F_TINY).pack(anchor="w", padx=14, pady=(12, 0))
        self.last_name = tk.Label(self.result_card, text=_("waiting", self._lang),
                                  bg=P.sidebar2, fg=P.text,
                                  font=("Segoe UI", 16, "bold"), anchor="w")
        self.last_name.pack(fill="x", padx=14, pady=(2, 0))
        self.last_detail = tk.Label(self.result_card, text="Start scanning to mark attendance.",
                                    bg=P.sidebar2, fg=P.muted,
                                    font=F_BODY, anchor="w")
        self.last_detail.pack(fill="x", padx=14, pady=(3, 12))

        # ── activity feed ─────────────────────────────────────────────────
        feed_hdr = tk.Frame(parent, bg=P.sidebar)
        feed_hdr.pack(fill="x", padx=22)
        tk.Label(feed_hdr, text=_("activity_feed", self._lang), bg=P.sidebar, fg=P.text,
                 font=F_H3).pack(side="left")

        feed_sh = tk.Frame(parent, bg="#070d16", highlightthickness=1,
                            highlightbackground=P.border_soft)
        feed_sh.pack(fill="both", expand=True, padx=22, pady=(8, 20))
        self.activity = tk.Listbox(feed_sh, bg="#070d16", fg="#c8d8ec",
                                   selectbackground=P.blue, relief="flat",
                                   highlightthickness=0, activestyle="none",
                                   font=F_MONO)
        sb = ttk.Scrollbar(feed_sh, orient="vertical", command=self.activity.yview)
        self.activity.configure(yscrollcommand=sb.set)
        self.activity.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        sb.pack(side="right", fill="y")
        self._feed_empty_hint = "  No activity yet \u2014 add a student or start scanning."
        self.activity.insert(0, self._feed_empty_hint)

    # ── right panel ───────────────────────────────────────────────────────────
    def _build_right_panel(self):
        self.dashboard_frame.grid_columnconfigure(0, weight=1)
        self.dashboard_frame.grid_rowconfigure(3, weight=1)   # row 3 = table (was row 4)

        # ── row 0 : title + subtitle ──────────────────────────────────────────
        title_bar = tk.Frame(self.dashboard_frame, bg=P.bg)
        title_bar.grid(row=0, column=0, sticky="ew", padx=24, pady=(10, 0))
        title_bar.grid_columnconfigure(0, weight=1)

        tk.Label(title_bar, text=_("teacher_dashboard", self._lang), bg=P.bg, fg=P.text,
                 font=F_H1).grid(row=0, column=0, sticky="w")
        self.date_lbl = tk.Label(title_bar, text=datetime.now().strftime("%A, %d %B %Y"),
                                 bg=P.bg, fg=P.muted, font=F_BODY)
        self.date_lbl.grid(row=1, column=0, sticky="w")

        # ── row 1 : action toolbar ────────────────────────────────────────────
        toolbar = tk.Frame(self.dashboard_frame, bg=P.surface, highlightthickness=1,
                            highlightbackground=P.border_soft)
        toolbar.grid(row=1, column=0, sticky="ew", padx=24, pady=(6, 0))

        def tbtn(parent, txt, cmd, bg, fg, abg, afg, anim=None):
            b = AnimBtn(parent, text=txt, command=cmd,
                bg=bg, fg=fg, activebackground=abg, activeforeground=afg,
                relief="flat", padx=14, pady=8, font=F_SMALL_B,
                cursor="hand2", anim_color=anim or abg)
            b.pack(side="left", padx=(0, 3), pady=8)
            return b

        grp = tk.Frame(toolbar, bg=P.surface)
        grp.pack(side="left", padx=8)
        # icons are already baked into these translation strings (e.g.
        # "➕  Add") — don't prefix them again here, that was the double
        # icon bug ("➕ ➕ Add").
        tbtn(grp, _("add_student", self._lang),    self.open_add_student,  P.green,    "white",   P.green_dark, "white",   "#15803d")
        tbtn(grp, _("edit_student", self._lang),   self.open_edit_student, P.amber,    "#1a1100", "#d97706",    "#1a1100", "#b45309")
        tbtn(grp, f'✅  {_("mark_attendance", self._lang)}', self.mark_attendance_selected, P.green_dark, "white", "#15803d", "white", "#0f6e2c")
        tbtn(grp, _("remove_student", self._lang), self.remove_selected,   P.red,      "white",   P.red_dark,   "white",   "#b91c1c")
        self.selection_count_lbl = tk.Label(grp, text="", bg=P.surface, fg=P.muted, font=F_SMALL)
        self.selection_count_lbl.pack(side="left", padx=(8, 0))
        tk.Frame(toolbar, bg=P.border, width=1).pack(side="left", fill="y", padx=8, pady=10)
        grp2 = tk.Frame(toolbar, bg=P.surface)
        grp2.pack(side="left")
        tbtn(grp2, _("import", self._lang), self._open_bulk_import, P.blue,     "white",   P.blue_dark,  "white",   P.blue_glow)
        tbtn(grp2, _("export", self._lang), self._export_menu,      P.sidebar2, P.ink,     P.border,     P.ink)

        # ── row 2 : stat cards ────────────────────────────────────────────────
        stats = tk.Frame(self.dashboard_frame, bg=P.bg)
        stats.grid(row=2, column=0, sticky="ew", padx=24, pady=(8, 6))
        for i in range(4):
            stats.grid_columnconfigure(i, weight=1, uniform="statcol")

        self.card_total   = StatCard(stats, _("total_students", self._lang), "0", "🎓", P.tint_blue,   P.blue,   subtitle=_("enrolled", self._lang))
        self.card_present = StatCard(stats, _("present_today", self._lang),  "0", "✓",  P.tint_green,  P.green,  subtitle=_("scanned_in", self._lang))
        self.card_absent  = StatCard(stats, _("absent_today", self._lang),   "0", "⏰", P.tint_amber,  P.amber,  subtitle=_("not_yet_seen", self._lang))
        self.card_xp      = StatCard(stats, _("xp_leader", self._lang),     "—", "★",  P.tint_purple, P.purple, subtitle=_("top_scorer", self._lang))
        for i, c in enumerate([self.card_total, self.card_present, self.card_absent, self.card_xp]):
            c.grid(row=0, column=i, sticky="nsew",
                   padx=(0 if i == 0 else 5, 0))

        # ── row 3 : search + table (combined frame) ───────────────────────────
        main_area = tk.Frame(self.dashboard_frame, bg=P.bg)
        main_area.grid(row=3, column=0, sticky="nsew", padx=24, pady=(0, 18))
        main_area.grid_rowconfigure(1, weight=1)
        main_area.grid_columnconfigure(0, weight=1)

        # search bar
        sb_frame = tk.Frame(main_area, bg=P.surface, highlightthickness=1,
                             highlightbackground=P.border)
        sb_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        sb_frame.grid_columnconfigure(3, weight=1)   # spacer column absorbs extra width

        tk.Label(sb_frame, text="🔍", bg=P.surface, fg=P.muted,
                 font=("Segoe UI", 12)).grid(row=0, column=0, padx=(14, 6), pady=10)
        # 3.0: no longer stretches to fill the whole row — a long, mostly-
        # empty search box was eating the space the Select All button
        # now uses right next to it.
        se = ttk.Entry(sb_frame, textvariable=self.search_var, font=("Segoe UI", 11), width=28)
        se.grid(row=0, column=1, sticky="w", pady=10)
        se.bind("<KeyRelease>", self._debounce_search)
        self.search_count_lbl = tk.Label(sb_frame, text="", bg=P.surface, fg=P.faint,
                                          font=F_SMALL)
        self.search_count_lbl.grid(row=0, column=2, padx=(8, 0))
        # column 3 is the flexible spacer (grid_columnconfigure above)
        AnimBtn(sb_frame, text=f'☑ {_("select_all", self._lang)}', command=self.select_all_students,
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=10, pady=6,
            font=F_SMALL_B, cursor="hand2").grid(row=0, column=4, padx=(0, 8), pady=8)
        AnimBtn(sb_frame, text=_("clear", self._lang), command=self._clear_search,
            bg=P.surface2, fg=P.muted, activebackground=P.border,
            activeforeground=P.text, relief="flat", padx=10, pady=6,
            font=F_SMALL_B, cursor="hand2").grid(row=0, column=5, padx=(0, 10), pady=8)

        # table
        tsh = tk.Frame(main_area, bg=P.surface, highlightthickness=1,
                        highlightbackground=P.border)
        tsh.grid(row=1, column=0, sticky="nsew")
        tsh.grid_rowconfigure(0, weight=1)
        tsh.grid_columnconfigure(0, weight=1)

        # 3.0: "extended" selectmode enables native Shift-click range-select
        # and Ctrl-click toggle-select, on top of the Select All button above.
        self.table = ttk.Treeview(tsh, columns=COLUMNS, show="tree headings",
                                   selectmode="extended")
        self.table.heading("#0", text="📷")
        self.table.column("#0", width=72, minwidth=60, stretch=False, anchor="center")

        col_w = {
            "Sl.No": 48, "Roll No": 80, "UID": 72, "Name": 160,
            "Grade": 64, "House": 72, "Parent Phone": 120,
            "Parent Email": 180, "No. of Days Present": 120,
            "XP": 62, "Multiplier": 80, "Streak": 68,
        }
        for col in COLUMNS:
            self.table.heading(col, text=col,
                               command=lambda c=col: self._sort_table(c))
            self.table.column(col, width=col_w.get(col, 80),
                              minwidth=44, anchor="center")
        self.table.column("Name",         anchor="w")
        self.table.column("Parent Email", anchor="w")
        self.table["displaycolumns"] = COLUMNS

        ys = ttk.Scrollbar(tsh, orient="vertical",   command=self.table.yview)
        xs = ttk.Scrollbar(tsh, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")

        self.table.tag_configure("odd",     background=P.surface)
        self.table.tag_configure("even",    background=P.surface2)
        self.table.tag_configure("present", background=P.green_soft)
        self.table.tag_configure("match",   background=P.blue_soft)
        self.table.tag_configure("hover",   background=P.hover)

        self.table.bind("<Double-1>", lambda _e: self.open_edit_student())
        self.table.bind("<Motion>",   self._on_table_hover)
        self.table.bind("<Leave>",    lambda _e: self._clear_table_hover())
        self.table.bind("<<TreeviewSelect>>", self._on_table_selection_changed)
        self._hovered_row    = None
        self._sort_col       = None
        self._sort_ascending = True

        # empty state
        self.table_empty = tk.Frame(tsh, bg=P.surface)
        tk.Label(self.table_empty, text="👥", bg=P.surface, fg=P.faint,
                 font=("Segoe UI", 32)).pack(pady=(50, 6))
        tk.Label(self.table_empty, text="No students enrolled yet",
                 bg=P.surface, fg=P.text, font=F_H3).pack()
        tk.Label(self.table_empty,
                 text='Click "➕ Add" in the toolbar to register your first student.',
                 bg=P.surface, fg=P.muted, font=F_BODY).pack(pady=(4, 0))

    def _on_table_hover(self, event):
        row_id = self.table.identify_row(event.y)
        if row_id == self._hovered_row:
            return
        self._clear_table_hover()
        if row_id:
            current = self.table.item(row_id, "tags")
            if "match" not in current and "present" not in current:
                self.table.item(row_id, tags=tuple(current) + ("hover",))
            self._hovered_row = row_id

    def _clear_table_hover(self):
        if self._hovered_row:
            current = self.table.item(self._hovered_row, "tags")
            self.table.item(self._hovered_row, tags=tuple(t for t in current if t != "hover"))
        self._hovered_row = None

    def _sort_table(self, col: str):
        if self._sort_col == col:
            self._sort_ascending = not self._sort_ascending
        else:
            self._sort_col = col
            self._sort_ascending = True
        self.refresh_table(self.search_var.get())


    # ── analytics tab ─────────────────────────────────────────────────────────
    def _build_analytics_tab(self):
        parent = self._tab_frames["Analytics"]
        parent.grid_rowconfigure(3, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        hdr = tk.Frame(parent, bg=P.sidebar)
        hdr.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        tk.Label(hdr, text="📊  Analytics", bg=P.sidebar, fg=P.text,
                 font=F_H2).pack(anchor="w")
        tk.Label(hdr, text="Attendance trends, top students, XP leaders",
                 bg=P.sidebar, fg=P.muted, font=F_SMALL).pack(anchor="w", pady=(2, 0))
        AnimBtn(hdr, text="↺ Refresh", command=self._refresh_analytics,
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=12, pady=6,
            font=F_SMALL_B, cursor="hand2").pack(anchor="e", pady=(6, 0))

        # mini stat cards
        sc = tk.Frame(parent, bg=P.sidebar)
        sc.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 6))
        sc.grid_columnconfigure(0, weight=1)
        sc.grid_columnconfigure(1, weight=1)
        self._an_card_rate = self._mini_stat(sc, 0, "Avg Attendance", "—%", P.tint_green, P.green)
        self._an_card_days = self._mini_stat(sc, 1, "School Days",    "0",  P.tint_blue,  P.blue)

        # house bar chart
        tk.Label(parent, text="House Attendance %", bg=P.sidebar,
                 fg=P.text, font=F_SMALL_B).grid(row=2, column=0, sticky="w", padx=18, pady=(4,2))
        self._chart_house_canvas = tk.Canvas(parent, bg=P.sidebar2, height=120,
                                              highlightthickness=1,
                                              highlightbackground=P.border_soft)
        self._chart_house_canvas.grid(row=2, column=0, sticky="ew", padx=18, pady=(20, 6))
        self._chart_house_canvas.bind("<Configure>", lambda _e: self._draw_house_chart())

        # trend line chart
        tk.Label(parent, text="30-Day Attendance Trend", bg=P.sidebar,
                 fg=P.text, font=F_SMALL_B).grid(row=3, column=0, sticky="nw", padx=18, pady=(4,2))
        self._chart_trend_canvas = tk.Canvas(parent, bg=P.sidebar2, height=120,
                                              highlightthickness=1,
                                              highlightbackground=P.border_soft)
        self._chart_trend_canvas.grid(row=3, column=0, sticky="new", padx=18, pady=(22, 6))
        self._chart_trend_canvas.bind("<Configure>", lambda _e: self._draw_trend_chart())

        # top 10 list
        tk.Label(parent, text="Top Students by XP", bg=P.sidebar,
                 fg=P.text, font=F_SMALL_B).grid(row=4, column=0, sticky="w", padx=18, pady=(4,2))
        top_sh = tk.Frame(parent, bg=P.sidebar2, highlightthickness=1,
                           highlightbackground=P.border_soft)
        top_sh.grid(row=5, column=0, sticky="nsew", padx=18, pady=(0, 10))
        top_sh.grid_rowconfigure(0, weight=1)
        top_sh.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(5, weight=1)
        self._an_top_list = tk.Listbox(top_sh, bg=P.sidebar2, fg=P.text,
                                        relief="flat", highlightthickness=0,
                                        font=F_BODY, activestyle="none",
                                        selectbackground=P.blue_soft,
                                        selectforeground=P.text)
        lsb = ttk.Scrollbar(top_sh, orient="vertical", command=self._an_top_list.yview)
        self._an_top_list.configure(yscrollcommand=lsb.set)
        self._an_top_list.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        lsb.grid(row=0, column=1, sticky="ns")

        # flame legend
        legend = tk.Frame(parent, bg=P.sidebar)
        legend.grid(row=6, column=0, sticky="ew", padx=18, pady=(0, 14))
        tk.Label(legend, text="🔥 5+ day streak  🔥🔥 10+  🔥🔥🔥 30+",
                 bg=P.sidebar, fg=P.faint, font=F_TINY).pack(anchor="w")

        # defer refresh until after window is fully drawn so canvas sizes are known
        parent.after(100, self._refresh_analytics)

    def _mini_stat(self, parent, col, title, value, tint, accent):
        f = tk.Frame(parent, bg=tint, highlightthickness=1, highlightbackground=accent)
        f.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 6, 0), pady=(0, 4))
        tk.Label(f, text=title, bg=tint, fg=P.muted, font=F_SMALL_B).pack(anchor="w", padx=12, pady=(10, 0))
        lbl = tk.Label(f, text=value, bg=tint, fg=P.text, font=("Segoe UI", 20, "bold"))
        lbl.pack(anchor="w", padx=12, pady=(2, 10))
        return lbl

    def _refresh_analytics(self):
        # Guard: called from refresh_all before analytics tab may be built
        if not hasattr(self, "_an_card_rate") or not hasattr(self, "_an_top_list"):
            return

        students = self.store.load_students()
        if not students:
            self._an_card_rate.configure(text="—%")
            self._an_card_days.configure(text="0")
            self._an_top_list.delete(0, "end")
            self._an_top_list.insert("end", "  No students enrolled yet.")
            if hasattr(self, "_chart_house_canvas"): self._draw_house_chart()
            if hasattr(self, "_chart_trend_canvas"): self._draw_trend_chart()
            return

        rows = self.store._load_audit_rows()
        unique_days = len({r[0] for r in rows if r}) or 1
        avg_pct = 0.0
        if students and unique_days:
            total_possible = len(students) * unique_days
            total_present  = sum(safe_int(s.get("No. of Days Present", 0)) for s in students)
            avg_pct = round(total_present / total_possible * 100, 1) if total_possible else 0.0
        self._an_card_rate.configure(text=f"{avg_pct}%")
        self._an_card_days.configure(text=str(unique_days))

        top = sorted(students, key=lambda s: safe_int(s.get("XP", 0)), reverse=True)[:10]
        self._an_top_list.delete(0, "end")
        medals = ["🥇", "🥈", "🥉"] + ["  "] * 20

        def flame(streak):
            s = safe_int(streak)
            if s >= 30: return "🔥🔥🔥"
            if s >= 10: return "🔥🔥"
            if s >= 5:  return "🔥"
            return ""

        for i, s in enumerate(top):
            xp     = safe_int(s.get("XP", 0))
            days   = safe_int(s.get("No. of Days Present", 0))
            streak = safe_int(s.get("Streak", 0))
            fl     = flame(streak)
            pct    = round(days / unique_days * 100, 1) if unique_days else 0.0
            house  = str(s.get("House", "")).strip()
            house_s = f"  [{house}]" if house and house != "None" else ""
            line = (f"  {medals[i]}  {s.get('Name','')}  ·  {s.get('Roll No','')} "
                    f"{house_s}  ·  {xp} XP  ·  {pct}% attendance  {fl}")
            self._an_top_list.insert("end", line)

        if hasattr(self, "_chart_house_canvas"): self._draw_house_chart()
        if hasattr(self, "_chart_trend_canvas"): self._draw_trend_chart()

    def _draw_house_chart(self):
        """Draw a simple bar chart of house attendance percentages."""
        if not hasattr(self, "_chart_house_canvas"): return
        canvas = self._chart_house_canvas
        canvas.delete("all")
        w = canvas.winfo_width()
        h = canvas.winfo_height()
        if w < 10 or h < 10:
            # canvas not yet laid out — the <Configure> binding will redraw
            # once it actually has a real size, so just bail quietly here.
            return
        stats = self.store.house_stats()
        houses = ["Red", "Blue", "Green", "Orange"]
        colors = ["#ef4444", "#3b82f6", "#22c55e", "#f97316"]
        values = [stats.get(h_, {}).get("pct", 0) for h_ in houses]
        max_v = max(values + [1])
        bar_w = max(1, (w - 40) // len(houses))
        for i, (house, val, color) in enumerate(zip(houses, values, colors)):
            bh = int((val / max_v) * (h - 40)) if max_v else 0
            x = 20 + i * bar_w + bar_w // 4
            canvas.create_rectangle(x, h - 20 - bh, x + bar_w // 2, h - 20,
                                     fill=color, outline="", tags="bar")
            canvas.create_text(x + bar_w // 4, h - 10, text=house, fill=P.text,
                               font=F_TINY, tags="label")
            canvas.create_text(x + bar_w // 4, h - 30 - bh, text=f"{val:.0f}%",
                               fill=P.text, font=F_TINY, tags="val")
        if all(v == 0 for v in values):
            canvas.create_text(w // 2, h // 2, text="No attendance recorded yet",
                               fill=P.muted, font=F_TINY)

    def _draw_trend_chart(self):
        """Draw a line chart of last 30 days attendance."""
        if not hasattr(self, "_chart_trend_canvas"): return
        canvas = self._chart_trend_canvas
        canvas.delete("all")
        w = canvas.winfo_width()
        h = canvas.winfo_height()
        if w < 10 or h < 10:
            return
        rows = self.store._load_audit_rows()
        from collections import Counter
        day_counts = Counter()
        for row in rows:
            if row and len(row) > 0:
                day_counts[row[0]] += 1
        # Get last 30 days
        dates = [(datetime.now() - __import__('datetime').timedelta(days=i)).strftime("%Y-%m-%d")
                 for i in range(29, -1, -1)]
        values = [day_counts.get(d, 0) for d in dates]
        max_v = max(values + [1])

        # Always draw the axis lines, even with zero data, so the box
        # never looks like a broken empty rectangle.
        canvas.create_line(20, h - 20, w - 20, h - 20, fill=P.chart_grid, width=1)
        canvas.create_line(20, 20, 20, h - 20, fill=P.chart_grid, width=1)

        if not any(values):
            canvas.create_text(w // 2, h // 2, text="No attendance recorded yet",
                               fill=P.muted, font=F_TINY)
            return

        points = []
        for i, v in enumerate(values):
            x = 20 + (i * (w - 40) // max(len(values) - 1, 1))
            y = h - 20 - int((v / max_v) * (h - 40))
            points.append((x, y))
        if len(points) > 1:
            for i in range(len(points) - 1):
                canvas.create_line(points[i][0], points[i][1], points[i+1][0], points[i+1][1],
                                   fill=P.blue, width=2)
        for x, y in points:
            canvas.create_oval(x-2, y-2, x+2, y+2, fill=P.blue, outline="")

    # ── bot tab ───────────────────────────────────────────────────────────────
    def _build_bot_tab(self):
        parent = self._tab_frames["Bot"]
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        hdr = tk.Frame(parent, bg=P.sidebar)
        hdr.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        tk.Label(hdr, text="🤖  AI Buddy", bg=P.sidebar, fg=P.text,
                 font=F_H2).pack(anchor="w")
        tk.Label(hdr, text="Ask about attendance, students, XP, or app actions.", bg=P.sidebar,
                 fg=P.muted, font=F_SMALL).pack(anchor="w", pady=(2, 0))

        chat_sh = tk.Frame(parent, bg="#070d16", highlightthickness=1, highlightbackground=P.border_soft)
        chat_sh.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        chat_sh.grid_rowconfigure(0, weight=1)
        chat_sh.grid_columnconfigure(0, weight=1)

        self.bot_chat = tk.Text(chat_sh, wrap="word", bg="#070d16", fg="#c8d8ec",
                                relief="flat", highlightthickness=0,
                                font=F_BODY, padx=12, pady=12, spacing3=6)
        self.bot_chat.tag_configure("speaker_you", foreground=P.blue, font=F_SMALL_B)
        self.bot_chat.tag_configure("speaker_bot", foreground=P.green, font=F_SMALL_B)
        bs = ttk.Scrollbar(chat_sh, orient="vertical", command=self.bot_chat.yview)
        self.bot_chat.configure(yscrollcommand=bs.set)
        self.bot_chat.grid(row=0, column=0, sticky="nsew")
        bs.grid(row=0, column=1, sticky="ns")
        self.bot_chat.insert("end", "AI Buddy: ", "speaker_bot")
        self.bot_chat.insert("end", "I use live school data. I will say when I do not know.\n")
        self.bot_chat.configure(state="disabled")

        inp = tk.Frame(parent, bg=P.sidebar)
        inp.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 18))
        inp.grid_columnconfigure(0, weight=1)
        self.bot_var = tk.StringVar()
        be = ttk.Entry(inp, textvariable=self.bot_var, font=F_BODY)
        be.grid(row=0, column=0, sticky="ew")
        be.bind("<Return>", lambda _: self._ask_bot())
        self.bot_send_btn = AnimBtn(inp, text="Ask",
            command=self._ask_bot, bg=P.blue, fg="white",
            activebackground=P.blue_dark, activeforeground="white",
            relief="flat", padx=14, pady=7,
            font=F_SMALL_B, cursor="hand2",
            anim_color=P.blue_glow)
        self.bot_send_btn.grid(row=0, column=1, padx=(8, 0))

    def _ask_bot(self):
        q = self.bot_var.get().strip()
        if not q or self.buddy_busy: return
        self.bot_var.set("")
        self._bot_append("You", q)
        self.buddy_busy = True
        self.bot_send_btn.configure(state="disabled", text="…")
        threading.Thread(target=self._run_bot, args=(q,), daemon=True).start()

    def _bot_append(self, speaker, msg):
        tag = "speaker_you" if speaker == "You" else "speaker_bot"
        self.bot_chat.configure(state="normal")
        self.bot_chat.insert("end", f"{speaker}: ", tag)
        self.bot_chat.insert("end", f"{msg}\n")
        self.bot_chat.see("end")
        self.bot_chat.configure(state="disabled")

    def _run_bot(self, q):
        try:    ans = self._bot_answer(q)
        except Exception as e: ans = f"Error: {e}"
        self.root.after(0, lambda: self._bot_append("AI Buddy", ans))
        self.root.after(0, lambda: (
            self.bot_send_btn.configure(state="normal", text="Ask"),
            setattr(self, "buddy_busy", False),
        ))

    def _bot_answer(self, q: str) -> str:
        """Answer the teacher's query. Uses keyword matching first (fast, always works),
        then falls back to Ollama only if the intent is unclear."""
        ql = q.lower().strip()

        # ── 3.0: bulk-action triggers (checked first — a multi-name command
        # like "remove these students - A, B, C" would otherwise get eaten
        # by the single-student remove/search matchers below) ───────────────
        bulk = self._parse_bulk_command(q, ql)
        if bulk is not None:
            return bulk

        # ── keyword pre-filter (handles ~90% of queries without Ollama) ──────
        # Search triggers
        if any(w in ql for w in ("search", "find", "look up", "show", "where is", "who is")):
            # extract the name/UID after the trigger word
            query = re.sub(r"^(search|find|look\s*up|show|where\s*is|who\s*is)\s*", "", ql).strip()
            if query:
                r = self._ai_search(query)
                n = r.get("count", 0)
                return f"Found {n} student(s) for '{query}'. Results shown in the table."
            return "What name or roll number should I search for?"

        # Remove/delete triggers
        if any(w in ql for w in ("remove", "delete", "kick", "expel")):
            query = re.sub(r"^(remove|delete|kick|expel)\s*(student)?\s*", "", ql).strip()
            if query:
                r = self._ai_remove(query)
                return str(r.get("message", "Done."))
            return "Which student should I remove? Give me their name or UID."

        if "export" in ql:
            self._run_on_ui(lambda: self._do_export("excel"))
            return "I started the styled Excel export."
        if "certificate" in ql:
            self._run_on_ui(self._open_certificates)
            return "The certificate generator is open. Choose the minimum attendance and generate."
        if any(w in ql for w in ("add student", "new student", "enrol", "enroll")):
            self._run_on_ui(self.open_add_student)
            return "The new-student form is open."
        if any(w in ql for w in ("start camera", "start scan", "start attendance")):
            self._run_on_ui(self.start_scan)
            return "I started the camera scan."
        if any(w in ql for w in ("stop camera", "stop scan")):
            self._run_on_ui(self.stop_scan)
            return "I stopped the camera scan."

        # Stats queries — answer from live data without Ollama
        if any(w in ql for w in ("xp", "points", "score")) and not any(w in ql for w in ("top", "leaderboard", "best")):
            students_all = self.store.load_students()
            matched = next((s for s in students_all if s.get("Name", "").lower() in ql or str(s.get("UID", "")).lower() in ql), None)
            if matched:
                return f"{matched.get('Name','')} has {safe_int(matched.get('XP',0))} XP."
            if any(w in ql for w in ("how many", "total", "all")):
                return f"The roster has {sum(safe_int(s.get('XP', 0)) for s in students_all)} XP in total across {len(students_all)} students."
            return "Tell me the student's name or UID so I can look up their recorded XP."

        if any(w in ql for w in ("how many", "total students", "count", "stats", "statistics")):
            students = self.store.load_students()
            present = len(self.scanned_today)
            absent  = len(students) - present
            top     = max(students, key=lambda s: safe_int(s.get("XP",0)), default=None)
            top_s   = f"{top.get('Name','')} ({top.get('XP',0)} XP)" if top else "N/A"
            return (f"Total students: {len(students)}\n"
                    f"Present today: {present}\n"
                    f"Absent today: {absent}\n"
                    f"XP leader: {top_s}")

        # XP query for specific student
        if any(w in ql for w in ("xp", "points", "score")) and not any(w in ql for w in ("top", "leaderboard", "best")):
            students_all = self.store.load_students()
            matched = next((s for s in students_all if s.get("Name","").lower() in ql or str(s.get("UID","")).lower() in ql), None)
            if matched:
                return (f"{matched.get('Name','')} has {safe_int(matched.get('XP',0))} XP "
                        f"({safe_int(matched.get('No. of Days Present',0))} days present, "
                        f"x{matched.get('Multiplier',1.0)} multiplier).")
            lines = [f"• {s.get('Name','')} — {safe_int(s.get('XP',0))} XP"
                     for s in sorted(students_all, key=lambda s: safe_int(s.get("XP",0)), reverse=True)]
            return ("XP for all students:\n" + "\n".join(lines)) if lines else "No students yet."

        # Top / leaderboard queries
        if any(w in ql for w in ("top", "leaderboard", "best", "highest xp", "most xp")):
            students = sorted(self.store.load_students(),
                              key=lambda s: safe_int(s.get("XP",0)), reverse=True)[:5]
            if not students: return "No students enrolled yet."
            lines = [f"{i+1}. {s.get('Name','')} — {s.get('XP',0)} XP"
                     for i, s in enumerate(students)]
            return "Top 5 by XP:\n" + "\n".join(lines)

        # Present/absent queries
        if "present" in ql or "attendance" in ql:
            students = self.store.load_students()
            present  = [s.get("Name","") for s in students if str(s.get("UID","")) in self.scanned_today]
            if not present: return "No students scanned yet today."
            return f"Present today ({len(present)}):\n" + "\n".join(f"• {n}" for n in present[:15])

        if "absent" in ql:
            students = self.store.load_students()
            absent   = [s.get("Name","") for s in students if str(s.get("UID","")) not in self.scanned_today]
            if not absent: return "Everyone is present!"
            return f"Absent today ({len(absent)}):\n" + "\n".join(f"• {n}" for n in absent[:15])

        # House queries
        if "house" in ql:
            stats = self.store.house_stats()
            ranked = sorted(stats.items(), key=lambda x: -x[1]["pct"])
            lines = [f"{h}: {d['pct']:.1f}% ({d['present']}/{d['students']})"
                     for h, d in ranked]
            return "House standings:\n" + "\n".join(lines)

        # Help
        if any(w in ql for w in ("help", "what can", "commands", "how do i")):
            return ("I can help you with:\n"
                    "• search <name/UID> — find a student\n"
                    "• remove <name/UID> — remove a student\n"
                    "• remove these students - A, B, C — remove several at once\n"
                    "• change these students to section A: X, Y, Z — bulk section change\n"
                    "• how many students — get stats\n"
                    "• top students — XP leaderboard\n"
                    "• who is present/absent today\n"
                    "• house standings\n"
                    "Or ask anything — I'll try Ollama if I'm unsure!")

        # ── Ollama fallback for unclear queries ───────────────────────────────
        return self._ollama_action_smart(q)

    def _ollama_action_smart(self, q: str) -> str:
        """Call Ollama with a robust prompt. Returns a plain-English answer."""
        # Build a concise context snapshot so the model can actually answer
        students = self.store.load_students()
        present  = len(self.scanned_today)
        present_uids  = self.scanned_today
        present_names = [s.get("Name","") for s in students if str(s.get("UID","")) in present_uids]
        absent_names  = [s.get("Name","") for s in students if str(s.get("UID","")) not in present_uids]
        student_list = "\n".join(
            f"  - {s.get('Name','')} | UID:{s.get('UID','')} | "
            f"Days:{s.get('No. of Days Present',0)} | XP:{s.get('XP',0)}"
            for s in students[:30]
        )
        context = (
            f"School: {self.settings.school_name or 'School'}\n"
            f"Total students: {len(students)}\n"
            f"Present today ({len(present_names)}): {', '.join(present_names) or 'none'}\n"
            f"Absent today ({len(absent_names)}): {', '.join(absent_names) or 'none'}\n"
            f"Roster:\n{student_list}\n"
        )
        if students:
            top = max(students, key=lambda s: safe_int(s.get("XP",0)), default=None)
            if top:
                context += f"XP leader: {top.get('Name','')} ({top.get('XP',0)} XP)\n"

        prompt = (
            f"{context}\n"
            f"You are FaceAttend AI — an assistant for school attendance.\n"
            f"RULES: Only use data from the roster above. Never invent numbers or names.\n"
            f"If a student is not in the roster, say so clearly.\n"
            f"Give the exact XP/days from the roster. Answer in 1-3 sentences. No JSON.\n\n"
            f"Teacher: {q}\n"
            f"Assistant:"
        )
        payload = {
            "model": OLLAMA_MODEL,
            "prompt": prompt,
            "stream": False,
            "options": {"temperature": 0.3, "num_predict": 120}
        }
        req = urllib.request.Request(
            f"{OLLAMA_URL.rstrip('/')}/api/generate",
            data=json.dumps(payload).encode(), method="POST",
            headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                data = json.loads(r.read())
            raw = str(data.get("response", "")).strip()
            return raw if raw else "I'm not sure how to answer that. Try: 'search <name>' or 'how many students'."
        except urllib.error.URLError:
            return ("Ollama isn't running — I answered what I could from the app data.\n"
                    "To enable full AI: install Ollama from ollama.ai, then run:\n"
                    f"  ollama pull {OLLAMA_MODEL}")
        except Exception as e:
            return f"AI error: {e}"

    # keep old _ollama_action for any legacy callers
    def _ollama_action(self, q) -> dict:
        ans = self._ollama_action_smart(q)
        return {"action": "answer", "query": "", "reply": ans}

    def _ai_search(self, query) -> dict:
        matches = self._find_students(query)
        def ui():
            self.search_var.set(query)
            self.refresh_table(query)
            self.log(f"AI: searched '{query}' → {len(matches)} result(s)")
        self._run_on_ui(ui)
        return {"ok": True, "count": len(matches)}

    def _ai_remove(self, query) -> dict:
        matches = self._find_students(query)
        if not matches: return {"ok": False, "message": "No student found."}
        if len(matches) > 1: return {"ok": False, "message": f"{len(matches)} matches. Be more specific."}
        s = matches[0]; uid = str(s.get("UID", "")); name = str(s.get("Name", uid))
        def ui():
            try:
                ok = messagebox.askyesno("AI Remove", f"Remove {name} ({uid})?")
                if ok:
                    play_sound(SOUND_NEGATIVE)
                    self.store.remove_student(uid)
                    self.scanned_today.discard(uid)
                    self._save_today_scans()
                    self.refresh_all()
                    self.log(f"AI removed {name} ({uid}).")
                return ok
            except Exception as e:
                self.log(f"AI remove UI error: {e}")
                return False
        removed = bool(self._run_on_ui(ui))
        return {"ok": removed, "message": "Removed." if removed else "Cancelled."}

    def _run_on_ui(self, fn):
        if threading.current_thread() is threading.main_thread():
            try:
                return fn()
            except Exception as e:
                self.log(f"UI fn error: {e}")
                raise
        ev  = threading.Event(); box: dict = {}
        def run():
            try: box["r"] = fn()
            except Exception as e: box["e"] = e
            finally: ev.set()
        self.root.after(0, run); ev.wait(timeout=60)
        if not ev.is_set():
            raise RuntimeError("UI operation timed out")
        if "e" in box: raise box["e"]
        return box.get("r")

    # ── 3.0: AI Buddy bulk actuation ─────────────────────────────────────────
    @staticmethod
    def _extract_name_tokens(text: str) -> list[str]:
        """
        Split a chunk like 'Aditi, Rohan and Meera' or 'Aditi; Rohan; Meera'
        into individual name/UID tokens, stripping the list punctuation a
        teacher would naturally type ("remove these students- A, B, C").
        """
        text = (text or "").strip(" -:\u2013\u2014")
        if not text:
            return []
        parts = re.split(r"\s*,\s*|\s+and\s+|\s*&\s*|\s*;\s*", text, flags=re.I)
        return [p.strip(" .") for p in parts if p.strip(" .")]

    def _parse_bulk_command(self, q_original: str, ql: str) -> str | None:
        """
        Recognises two multi-student AI Buddy commands and executes them:
          • "remove these students - A, B, C" / "delete students: A, B"
          • "change these students to section A: X, Y, Z" /
            "move X, Y to section 8B" / "change section of X, Y to grade 8B"
        Returns the reply string if a bulk command was matched and handled
        (even if it failed/needed clarification), or None if q wasn't a
        bulk command — in which case _bot_answer() continues to its normal
        single-student and Q&A matchers.
        """
        # ── bulk section/grade change ────────────────────────────────────────
        sec_m = re.search(r"\bto\s+(?:section|grade)\s+([a-z0-9]{1,4})\b", ql, re.I)
        change_words = ("change", "move", "reassign", "shift")
        if sec_m and any(w in ql for w in change_words) and "student" in ql:
            target = sec_m.group(1).upper()
            remainder = (ql[:sec_m.start()] + ql[sec_m.end():])
            remainder = re.sub(r"^(change|move|reassign|shift)\s+(the\s+|these\s+)?students?\s*"
                                r"(of|to)?\s*", "", remainder).strip()
            remainder = re.sub(r"^section\s+of\s*", "", remainder).strip()
            names = self._extract_name_tokens(remainder)
            if len(names) >= 1:
                return self._ai_bulk_section_change(names, target)

        # ── bulk remove ───────────────────────────────────────────────────────
        rm_m = re.match(r"^(?:remove|delete|kick|expel)\s+(?:these\s+|the\s+)?students?\s*"
                         r"(?:[-:\u2013\u2014]|from\s+the\s+roster)?\s*(.+)$", ql, re.I)
        if rm_m:
            names = self._extract_name_tokens(rm_m.group(1))
            if len(names) >= 2:
                return self._ai_bulk_remove(names)

        return None

    def _resolve_bulk_tokens(self, tokens: list[str]) -> tuple[list[dict], list[str]]:
        """Resolve each free-text token to exactly one student. Returns
        (resolved_students, problem_tokens) — a token is a "problem" if it
        matched zero students or more than one (teacher should be more
        specific for that one, e.g. use the UID instead of a common name)."""
        resolved, problems = [], []
        for tok in tokens:
            matches = self._find_students(tok)
            if len(matches) == 1:
                resolved.append(matches[0])
            else:
                problems.append(tok)
        return resolved, problems

    def _ai_bulk_remove(self, tokens: list[str]) -> str:
        resolved, problems = self._resolve_bulk_tokens(tokens)
        if not resolved:
            return (f"I couldn't uniquely match any of: {', '.join(tokens)}. "
                     "Try their UID or a more specific name.")

        def ui():
            names = [str(s.get("Name", "")) for s in resolved]
            uids  = [str(s.get("UID", "")) for s in resolved]
            preview = ", ".join(names[:8]) + ("" if len(names) <= 8 else f", +{len(names)-8} more")
            if not messagebox.askyesno("AI Bulk Remove",
                                        f"Remove {len(resolved)} student(s)?\n\n{preview}"):
                return False
            play_sound(SOUND_NEGATIVE)
            self.store.bulk_remove(uids)
            for uid in uids:
                self.scanned_today.discard(uid)
            self._save_today_scans()
            self.refresh_all()
            self.log(f"AI bulk-removed {len(uids)} student(s): {', '.join(names)}.")
            return True

        ok = bool(self._run_on_ui(ui))
        if not ok:
            return "Cancelled — nothing was removed."
        msg = f"Removed {len(resolved)} student(s)."
        if problems:
            msg += f" Couldn't match: {', '.join(problems)} — try their UID instead."
        return msg

    def _ai_bulk_section_change(self, tokens: list[str], target_grade: str) -> str:
        resolved, problems = self._resolve_bulk_tokens(tokens)
        if not resolved:
            return (f"I couldn't uniquely match any of: {', '.join(tokens)}. "
                     "Try their UID or a more specific name.")

        def ui():
            names = [str(s.get("Name", "")) for s in resolved]
            uids  = [str(s.get("UID", "")) for s in resolved]
            preview = ", ".join(names[:8]) + ("" if len(names) <= 8 else f", +{len(names)-8} more")
            if not messagebox.askyesno("AI Bulk Section Change",
                                        f"Move {len(resolved)} student(s) to {target_grade}?\n\n{preview}"):
                return False
            self.store.bulk_set_grades({uid: target_grade for uid in uids})
            self.students_by_uid = self.store.students_by_uid()
            self.refresh_all()
            play_sound(SOUND_SUCCESS)
            self.log(f"AI bulk section-change: {', '.join(names)} → {target_grade}.")
            return True

        ok = bool(self._run_on_ui(ui))
        if not ok:
            return "Cancelled — no sections were changed."
        msg = f"Moved {len(resolved)} student(s) to {target_grade}."
        if problems:
            msg += f" Couldn't match: {', '.join(problems)} — try their UID instead."
        return msg

    def _find_students(self, query) -> list[dict]:
        q = query.lower()
        return [s for s in self.store.load_students()
                if q in " ".join(str(v) for v in s.values()).lower()]

    # ── mail tab ──────────────────────────────────────────────────────────────
    def _build_mail_tab(self):
        parent = self._tab_frames["Mail"]
        parent.grid_rowconfigure(4, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        hdr = tk.Frame(parent, bg=P.sidebar)
        hdr.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        tk.Label(hdr, text="✉  Messages", bg=P.sidebar, fg=P.text,
                 font=F_H2).pack(anchor="w")
        tk.Label(hdr, text="Email a parent directly", bg=P.sidebar,
                 fg=P.muted, font=F_SMALL).pack(anchor="w", pady=(2, 0))

        def card(row):
            f = tk.Frame(parent, bg=P.sidebar2, highlightthickness=1, highlightbackground=P.border)
            f.grid(row=row, column=0, sticky="ew", padx=16, pady=(0, 10))
            f.grid_columnconfigure(1, weight=1)
            return f

        # student picker
        pc = card(1)
        tk.Label(pc, text="Student", bg=P.sidebar2, fg=P.muted,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=14, pady=(12,4))
        self._msg_stu_var = tk.StringVar()
        self._msg_stu_cb  = ttk.Combobox(pc, textvariable=self._msg_stu_var,
                                          font=("Segoe UI", 10), state="normal")
        self._msg_stu_cb.grid(row=0, column=1, sticky="ew", padx=14, pady=(12,4))
        self._msg_stu_cb.bind("<<ComboboxSelected>>", self._msg_student_pick)
        self._msg_stu_cb.bind("<Return>", self._msg_student_pick)
        AnimBtn(pc, text="↻", command=self._msg_refresh_list,
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=8, pady=2,
            font=("Segoe UI", 9, "bold"), cursor="hand2").grid(row=0, column=2, padx=(0,14), pady=(12,4))
        tk.Label(pc, text="Type name / roll / UID then press Enter, or pick from list",
                 bg=P.sidebar2, fg=P.muted,
                 font=("Segoe UI", 8)).grid(row=1, column=0, columnspan=3, sticky="w", padx=14, pady=(0,10))

        # info
        ic = card(2)
        def irow(txt, r, var, ro=False):
            tk.Label(ic, text=txt, bg=P.sidebar2, fg=P.muted,
                                          font=("Segoe UI", 9, "bold")).grid(row=r, column=0, sticky="w", padx=14, pady=6)
            e = ttk.Entry(ic, textvariable=var, font=("Segoe UI", 10))
            e.grid(row=r, column=1, sticky="ew", padx=14, pady=6)
            if ro: e.configure(state="readonly")
            return e
        self._msg_from_var    = tk.StringVar(value="(set in Settings)")
        self._msg_to_email    = tk.StringVar()
        self._msg_to_phone    = tk.StringVar()
        irow("From (email)", 0, self._msg_from_var, ro=True)
        irow("Parent email", 1, self._msg_to_email)

        # subject
        sf = tk.Frame(parent, bg=P.sidebar)
        sf.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 6))
        sf.grid_columnconfigure(1, weight=1)
        tk.Label(sf, text="Subject", bg=P.sidebar, fg=P.muted,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0,8))
        self._msg_subject = tk.StringVar(value="Message from School")
        ttk.Entry(sf, textvariable=self._msg_subject, font=("Segoe UI", 10)).grid(row=0, column=1, sticky="ew")

        # body
        bsh = tk.Frame(parent, bg="#070d16", highlightthickness=1, highlightbackground=P.border)
        bsh.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 8))
        bsh.grid_rowconfigure(0, weight=1)
        bsh.grid_columnconfigure(0, weight=1)
        self._msg_body = tk.Text(bsh, wrap="word", bg="#070d16", fg="#c8d8ec",
                                 relief="flat", highlightthickness=0,
                                 font=("Segoe UI", 10), padx=10, pady=10)
        ms = ttk.Scrollbar(bsh, orient="vertical", command=self._msg_body.yview)
        self._msg_body.configure(yscrollcommand=ms.set)
        self._msg_body.grid(row=0, column=0, sticky="nsew")
        ms.grid(row=0, column=1, sticky="ns")
        self._msg_body.insert("1.0", "")

        # buttons
        bf = tk.Frame(parent, bg=P.sidebar)
        bf.grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 16))
        AnimBtn(bf, text="📧  Send Email", command=self._send_email,
            bg=P.blue, fg="white", activebackground=P.blue_dark,
            activeforeground="white", relief="flat", padx=14, pady=9,
            font=("Segoe UI", 10, "bold"), cursor="hand2",
            anim_color=P.blue_glow).pack(side="left", padx=(0, 8))
        # WhatsApp button removed in v2.9
        self._msg_refresh_list()

    def _msg_refresh_list(self):
        students = self.store.load_students()
        self._msg_stu_map = {}
        opts = []
        for s in students:
            lbl = f"{s.get('Name','')}  ({s.get('Roll No','')} / {s.get('UID','')})"
            opts.append(lbl); self._msg_stu_map[lbl] = s
        if hasattr(self, "_msg_stu_cb"):
            self._msg_stu_cb.configure(values=opts)

    def _msg_student_pick(self, _e=None):
        typed = self._msg_stu_var.get().strip()
        s = self._msg_stu_map.get(typed)
        if s is None:
            m = self._find_students(typed); s = m[0] if m else None
        if s is None: return
        lbl = f"{s.get('Name','')}  ({s.get('Roll No','')} / {s.get('UID','')})"
        self._msg_stu_var.set(lbl)
        self._msg_to_email.set(str(s.get("Parent Email", "")))
        self._msg_to_phone.set(str(s.get("Parent Phone", "")))

    def _send_email(self):
        to    = self._msg_to_email.get().strip()
        subj  = self._msg_subject.get().strip() or "Message from School"
        body  = self._msg_body.get("1.0", "end").strip()
        sender= self.settings.smtp_sender or ""
        pw    = self.settings.smtp_password or ""
        if not to:     play_sound(SOUND_ERROR); Toast.show(self.root, "Select a student first.", "warning"); return
        if not body:   play_sound(SOUND_ERROR); Toast.show(self.root, "Message body is empty.", "warning"); return
        if not sender: play_sound(SOUND_ERROR); Toast.show(self.root, "Set sender email in Settings.", "error"); return
        def worker():
            try:
                import smtplib
                from email.message import EmailMessage
                em = EmailMessage()
                em["Subject"] = subj; em["From"] = sender; em["To"] = to
                em.set_content(body)
                with smtplib.SMTP("smtp.gmail.com", 587) as srv:
                    srv.starttls(); srv.login(sender, pw); srv.send_message(em)
                self.root.after(0, lambda: Toast.show(self.root, f"Email sent to {to}", "success"))
                self.root.after(0, lambda: self.log(f"Email sent to {to}"))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda: (
                    play_sound(SOUND_ERROR),
                    Toast.show(self.root, f"Email failed: {err}", "error"),
                ))
        threading.Thread(target=worker, daemon=True).start()


    def _attendance_message(self, name: str, roll: str, status: str, xp: int, now: datetime) -> str:
        school   = (self.settings.school_name or "FaceAttend").strip()
        template = (self.settings.auto_message_template or "").strip()
        if not template:
            template = Settings._defaults["auto_message_template"]
        return template.format(
            name   = name,
            roll   = roll,
            status = status,
            time   = now.strftime("%I:%M %p"),
            date   = now.strftime("%d %B %Y"),
            xp     = xp,
            school = school,
        )

    def _absent_message(self, name: str, roll: str, now: datetime) -> str:
        school   = (self.settings.school_name or "FaceAttend").strip()
        template = (self.settings.absent_message_template or "").strip()
        if not template:
            template = Settings._defaults["absent_message_template"]
        return template.format(
            name=name, roll=roll,
            time=now.strftime("%I:%M %p"), date=now.strftime("%d %B %Y"),
            school=school,
        )

    def get_absent_students(self) -> list[dict]:
        """Students not yet marked present today, only on active school days."""
        if not self.schedule.is_active():
            return []
        students = self.store.load_students()
        return [s for s in students if str(s.get("UID", "")) not in self.scanned_today]

    def send_absent_notifications(self, manual: bool = False) -> dict:
        """
        Sends an absence notice to every student not yet marked present
        today, via email and/or WhatsApp depending on Settings. Shared by
        both the fixed-time daily scheduler and the manual 'Send to All
        Absent Parents Now' button / console command.
        Returns a small result dict for logging/toasting.
        """
        absentees = self.get_absent_students()
        if not absentees:
            return {"sent_email": 0, "sent_whatsapp": 0, "skipped": 0, "total": 0}

        now = datetime.now()
        send_email    = self.settings.absent_notify_email
        send_whatsapp = False  # WhatsApp removed
        sender = (self.settings.smtp_sender or "").strip()
        pw     = (self.settings.smtp_password or "").strip()

        sent_email = sent_whatsapp = skipped = 0

        for s in absentees:
            name  = str(s.get("Name", ""))
            roll  = str(s.get("Roll No", ""))
            email = str(s.get("Parent Email", "")).strip()
            phone = str(s.get("Parent Phone", "")).strip()
            body  = self._absent_message(name, roll, now)

            if send_email:
                if email and sender and pw:
                    def worker(to=email, n=name, msg_body=body, subj=f"Attendance: {name} is absent today"):
                        try:
                            import smtplib
                            from email.message import EmailMessage
                            em = EmailMessage()
                            em["Subject"] = subj
                            em["From"]    = sender
                            em["To"]      = to
                            em.set_content(msg_body)
                            with smtplib.SMTP("smtp.gmail.com", 587) as srv:
                                srv.starttls(); srv.login(sender, pw); srv.send_message(em)
                            self.root.after(0, lambda: self.log(f"✅ Absence email sent to {to} ({n})."))
                        except Exception as e:
                            err = str(e)
                            self.root.after(0, lambda: self.log(f"❌ Absence email failed ({n}): {err}"))
                    threading.Thread(target=worker, daemon=True).start()
                    sent_email += 1
                else:
                    skipped += 1

            if send_whatsapp:
                if phone:
                    pass  # WhatsApp removed
                    sent_whatsapp += 1
                else:
                    skipped += 1

        result = {"sent_email": sent_email, "sent_whatsapp": sent_whatsapp,
                   "skipped": skipped, "total": len(absentees)}
        self.log(f"Absentee notify ({'manual' if manual else 'scheduled'}): "
                 f"{result['sent_email']} email(s) "
                 f"{result['skipped']} skipped, {result['total']} absent total.")
        return result

    def _check_absent_notify_schedule(self):
        """
        Runs every minute on the UI thread (piggybacks on the existing
        clock tick). Fires send_absent_notifications() once per day, the
        first time 'now' reaches the configured absent_notify_time —
        guarded by a date stamp in Settings so it never double-fires.
        """
        if not self.settings.absent_notify_enabled:
            return
        now = datetime.now()
        try:
            h, m = str(self.settings.absent_notify_time or "09:30").split(":")
            target = now.replace(hour=int(h), minute=int(m), second=0, microsecond=0)
        except Exception:
            return
        already_ran = self.settings._data.get("_absent_notify_last_run_date") == now.strftime("%Y-%m-%d")
        if now >= target and not already_ran:
            self.settings.set("_absent_notify_last_run_date", now.strftime("%Y-%m-%d"))
            self.settings.save()
            result = self.send_absent_notifications(manual=False)
            if result["total"] > 0:
                play_sound(SOUND_EXPORT)
                Toast.show(self.root,
                           f"Scheduled absentee alert sent: {result['total']} student(s) absent.",
                           "info", 6000)


    def _auto_notify_parent(self, student: dict, status: str, xp: int, now: datetime) -> None:
        """Fired once per student per day, right after attendance is recorded."""
        name  = str(student.get("Name", ""))
        roll  = str(student.get("Roll No", ""))
        email = str(student.get("Parent Email", "")).strip()
        phone = str(student.get("Parent Phone", "")).strip()
        body  = self._attendance_message(name, roll, status, xp, now)

        if self.settings.auto_notify_email:
            sender = (self.settings.smtp_sender or "").strip()
            pw     = (self.settings.smtp_password or "").strip()
            if not email:
                msg = f"Auto-email skipped ({name}): no parent email on file."
                self.log(msg); Toast.show(self.root, msg, "warning")
            elif not sender or not pw:
                msg = "Auto-email skipped: add Gmail address & app password in Settings → Email Sender."
                self.log(msg); Toast.show(self.root, msg, "warning")
            else:
                self.log(f"Sending email to {email} ({name})…")
                def worker(to=email, subj=f"Attendance: {name} is {status}",
                           msg_body=body, s=sender, p=pw, n=name):
                    try:
                        import smtplib
                        from email.message import EmailMessage
                        em = EmailMessage()
                        em["Subject"] = subj
                        em["From"]    = s
                        em["To"]      = to
                        em.set_content(msg_body)
                        with smtplib.SMTP("smtp.gmail.com", 587) as srv:
                            srv.starttls()
                            srv.login(s, p)
                            srv.send_message(em)
                        self.root.after(0, lambda: (
                            self.log(f"✅ Email sent to {to} ({n})."),
                            Toast.show(self.root, f"Email sent to parent of {n}", "success"),
                        ))
                    except Exception as e:
                        err = str(e)
                        self.root.after(0, lambda: (
                            self.log(f"❌ Email failed ({n}): {err}"),
                            play_sound(SOUND_ERROR),
                            Toast.show(self.root, f"Email failed: {err[:60]}", "error"),
                        ))
                threading.Thread(target=worker, daemon=True).start()

        if self.settings.auto_notify_whatsapp:
            if not phone:
                msg = f"Auto-WhatsApp skipped ({name}): no parent phone on file."
                self.log(msg); Toast.show(self.root, msg, "warning")
            else:
                pass  # WhatsApp removed
                self.log(f"📱 WhatsApp opened for {name} ({phone}).")
                Toast.show(self.root, f"WhatsApp opened for parent of {name}", "info")


    # ── house competition tab ─────────────────────────────────────────────────
    def _build_houses_tab(self):
        parent = self._tab_frames["Houses"]
        parent.grid_rowconfigure(2, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # header
        hdr = tk.Frame(parent, bg=P.sidebar)
        hdr.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        hdr.grid_columnconfigure(0, weight=1)
        tk.Label(hdr, text="🏠  House Competition", bg=P.sidebar, fg=P.text,
                 font=F_H2).grid(row=0, column=0, sticky="w")
        tk.Label(hdr, text="Live attendance % by house · auto-updates on every scan",
                 bg=P.sidebar, fg=P.muted, font=F_SMALL).grid(row=1, column=0, sticky="w")
        AnimBtn(hdr, text="↺", command=self._refresh_houses,
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=10, pady=4,
            font=F_BODY_B, cursor="hand2").grid(row=0, column=1, rowspan=2, sticky="e")

        # 2×2 grid of house cards — prevents truncation no matter sidebar width
        cards_frame = tk.Frame(parent, bg=P.sidebar)
        cards_frame.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
        cards_frame.grid_columnconfigure(0, weight=1)
        cards_frame.grid_columnconfigure(1, weight=1)

        self._house_cards = {}
        house_display = [
            ("Red",    "#ef4444", "#2a1316"),
            ("Blue",   "#3b82f6", "#132236"),
            ("Green",  "#22c55e", "#0d2218"),
            ("Orange", "#f97316", "#271608"),
        ]
        positions = [(0, 0), (0, 1), (1, 0), (1, 1)]
        for idx, (house, accent, tint) in enumerate(house_display):
            r, c = positions[idx]
            card = tk.Frame(cards_frame, bg=tint, highlightthickness=2,
                            highlightbackground=accent)
            card.grid(row=r, column=c, sticky="ew",
                      padx=(0 if c == 0 else 6, 0),
                      pady=(0 if r == 0 else 6, 0))

            top_row = tk.Frame(card, bg=tint)
            top_row.pack(fill="x", padx=12, pady=(12, 0))
            tk.Label(top_row, text="●", bg=tint, fg=accent,
                     font=("Segoe UI", 14)).pack(side="left")
            tk.Label(top_row, text=f"  {house}", bg=tint, fg=P.text,
                     font=F_H3).pack(side="left")

            pct_lbl = tk.Label(card, text="0.0%", bg=tint, fg=accent,
                               font=("Segoe UI", 26, "bold"))
            pct_lbl.pack(anchor="w", padx=14, pady=(4, 0))

            det_lbl = tk.Label(card, text="0 / 0 present", bg=tint,
                               fg=P.muted, font=F_SMALL)
            det_lbl.pack(anchor="w", padx=14, pady=(0, 12))

            self._house_cards[house] = (pct_lbl, det_lbl)

        # leaderboard — Listbox so lines never wrap
        lb_hdr = tk.Frame(parent, bg=P.sidebar)
        lb_hdr.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))
        lb_hdr.grid_rowconfigure(1, weight=1)
        lb_hdr.grid_columnconfigure(0, weight=1)

        tk.Label(lb_hdr, text="RANKING", bg=P.sidebar, fg=P.muted,
                 font=F_SMALL_B).grid(row=0, column=0, sticky="w", pady=(0, 6))

        lb_sh = tk.Frame(lb_hdr, bg=P.sidebar2, highlightthickness=1,
                          highlightbackground=P.border_soft)
        lb_sh.grid(row=1, column=0, sticky="nsew")
        lb_sh.grid_rowconfigure(0, weight=1)
        lb_sh.grid_columnconfigure(0, weight=1)

        self.house_lb = tk.Listbox(lb_sh, bg=P.sidebar2, fg=P.text,
                                    relief="flat", highlightthickness=0,
                                    font=F_BODY, activestyle="none",
                                    selectbackground=P.hover,
                                    selectforeground=P.text,
                                    height=5)
        lsb = ttk.Scrollbar(lb_sh, orient="vertical", command=self.house_lb.yview)
        self.house_lb.configure(yscrollcommand=lsb.set)
        self.house_lb.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        lsb.grid(row=0, column=1, sticky="ns")

        self._refresh_houses()

    def _refresh_houses(self):
        stats = self.store.house_stats()
        ranked = sorted(stats.items(), key=lambda x: -x[1]["pct"])
        medals = ["🥇 1st", "🥈 2nd", "🥉 3rd", "   4th"]

        for house, (pct_lbl, det_lbl) in self._house_cards.items():
            d = stats[house]
            pct_lbl.configure(text=f"{d['pct']:.1f}%")
            det_lbl.configure(text=f"{d['present']} / {d['students']} present")

        self.house_lb.delete(0, "end")
        for rank, (house, d) in enumerate(ranked):
            pct   = f"{d['pct']:.1f}%"
            entry = f"  {medals[rank]}   {house:<8}   {pct:>6}   ({d['present']}/{d['students']} students)"
            self.house_lb.insert("end", entry)

    # ── bulk import dialog ─────────────────────────────────────────────────────
    def _open_bulk_import(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Select students.xlsx to import",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            parent=self.root,
        )
        if not path: return
        path = Path(path)

        dlg = tk.Toplevel(self.root)
        dlg.title("Bulk Import")
        dlg.geometry("500x320")
        dlg.configure(bg=P.bg)
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="📥  Bulk Import", bg=P.bg, fg=P.text, font=F_H2).pack(anchor="w", padx=24, pady=(20,4))
        tk.Label(dlg, text=f"Importing from: {path.name}", bg=P.bg, fg=P.muted, font=F_SMALL).pack(anchor="w", padx=24)
        tk.Label(dlg, text="Required columns: Roll No, Name, Grade\nOptional: House, Parent Phone, Parent Email",
                 bg=P.bg, fg=P.faint, font=F_TINY, justify="left").pack(anchor="w", padx=24, pady=(4,12))

        log_box = tk.Text(dlg, bg=P.surface, fg=P.text, relief="flat",
                          font=F_MONO, height=7, padx=10, pady=8, highlightthickness=0)
        log_box.pack(fill="x", padx=24)
        log_box.insert("end", "Ready. Click Import to begin.\n")
        log_box.configure(state="disabled")

        def do_import():
            log_box.configure(state="normal")
            log_box.delete("1.0", "end")
            log_box.insert("end", "Importing...\n")
            dlg.update()
            try:
                added, skipped = self.store.bulk_import(path)
                log_box.insert("end", f"✅  Added {added} student(s).\n")
                for s in skipped:
                    log_box.insert("end", f"⚠  {s}\n")
                self.refresh_all()
                self.log(f"Bulk import: {added} added, {len(skipped)} skipped.")
                play_sound(SOUND_EXPORT)
                Toast.show(self.root, f"Imported {added} students.", "success")
            except Exception as e:
                log_box.insert("end", f"❌  Error: {e}\n")
                play_sound(SOUND_ERROR)
            log_box.configure(state="disabled")

        btn_row = tk.Frame(dlg, bg=P.bg)
        btn_row.pack(fill="x", padx=24, pady=12)
        AnimBtn(btn_row, text="📥  Import", command=do_import,
            bg=P.blue, fg="white", activebackground=P.blue_dark,
            activeforeground="white", relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color=P.blue_glow).pack(side="left")
        AnimBtn(btn_row, text="✕  Close", command=dlg.destroy,
            bg=P.surface2, fg=P.text, activebackground=P.border,
            activeforeground=P.text, relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2").pack(side="left", padx=(10, 0))

    # ── attendance certificate generator ──────────────────────────────────────
    def _open_certificates(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Attendance Certificate Generator")
        dlg.geometry("520x420")
        dlg.configure(bg=P.bg)
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="🏅  Attendance Certificates", bg=P.bg, fg=P.text, font=F_H2).pack(anchor="w", padx=24, pady=(20,4))
        tk.Label(dlg, text="Generates a PDF certificate per eligible student", bg=P.bg, fg=P.muted, font=F_SMALL).pack(anchor="w", padx=24)

        row = tk.Frame(dlg, bg=P.bg)
        row.pack(anchor="w", padx=24, pady=(12,4))
        tk.Label(row, text="Minimum attendance %:", bg=P.bg, fg=P.text, font=F_BODY_B).pack(side="left")
        pct_var = tk.StringVar(value="90")
        ttk.Entry(row, textvariable=pct_var, font=F_BODY, width=6).pack(side="left", padx=8)

        log_box = tk.Text(dlg, bg=P.surface, fg=P.text, relief="flat",
                          font=F_MONO, height=8, padx=10, pady=8, highlightthickness=0)
        log_box.pack(fill="x", padx=24, pady=(8, 0))
        log_box.insert("end", "Click Generate to build certificates.\n")
        log_box.configure(state="disabled")

        def do_generate():
            try:
                min_pct = float(pct_var.get())
            except ValueError:
                play_sound(SOUND_ERROR)
                Toast.show(self.root, "Enter a valid percentage.", "warning"); return

            eligible = self.store.generate_certificates(min_pct)
            log_box.configure(state="normal")
            log_box.delete("1.0", "end")
            if not eligible:
                log_box.insert("end", f"No students meet {min_pct}% threshold.\n")
                log_box.configure(state="disabled"); return

            EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
            generated = 0
            for s in eligible:
                name   = str(s.get("Name",""))
                roll   = str(s.get("Roll No",""))
                grade  = str(s.get("Grade",""))
                pct    = s["_pct"]
                days   = safe_int(s.get("No. of Days Present",0))
                school = self.settings.school_name.strip() or "School"
                fname  = EXPORTS_DIR / f"Certificate_{roll}_{name.replace(' ','_')}.pdf"
                ok = self._generate_certificate_pdf(fname, name, roll, grade, pct, days, school)
                if ok:
                    log_box.insert("end", f"✅  {name} ({roll}) — {pct}% — saved\n")
                    generated += 1
                else:
                    log_box.insert("end", f"⚠  {name} — PDF generation failed (install reportlab)\n")
            log_box.configure(state="disabled")
            Toast.show(self.root, f"{generated} certificate(s) saved to exports/", "success")
            self.log(f"Generated {generated} attendance certificates.")

        btn_row = tk.Frame(dlg, bg=P.bg)
        btn_row.pack(fill="x", padx=24, pady=10)
        AnimBtn(btn_row, text="🏅  Generate", command=do_generate,
            bg=P.amber, fg="#1a1100", activebackground="#d97706",
            activeforeground="#1a1100", relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color="#b45309").pack(side="left")
        AnimBtn(btn_row, text="📂  Open Folder", command=lambda: self._open_exports_folder(),
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2").pack(side="left", padx=(10,0))
        AnimBtn(btn_row, text="✕  Close", command=dlg.destroy,
            bg=P.surface2, fg=P.text, activebackground=P.border,
            activeforeground=P.text, relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2").pack(side="right")

    def _generate_certificate_pdf(self, path: Path, name, roll, grade, pct, days, school) -> bool:
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib import colors
            from reportlab.pdfgen import canvas as rl_canvas
            import math
        except ImportError:
            return False

        page_w, page_h = landscape(A4)
        navy   = colors.HexColor("#1e3a5f")
        gold   = colors.HexColor("#c9a227")
        cream  = colors.HexColor("#faf6ee")
        ink    = colors.HexColor("#1a1a2e")
        muted  = colors.HexColor("#5c6370")
        accent = colors.HexColor("#2563eb")

        c = rl_canvas.Canvas(str(path), pagesize=landscape(A4))
        cx = page_w / 2

        # Background
        c.setFillColor(cream)
        c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

        # Decorative borders
        for inset, stroke, color in ((22, 3.5, navy), (30, 1.5, gold), (38, 0.75, navy)):
            c.setStrokeColor(color)
            c.setLineWidth(stroke)
            c.roundRect(inset, inset, page_w - 2 * inset, page_h - 2 * inset,
                        6, fill=0, stroke=1)

        # Corner ornaments
        def corner_ornament(ox, oy, sx, sy):
            c.setStrokeColor(gold)
            c.setLineWidth(1.2)
            for i in range(4):
                c.line(ox, oy + i * 8 * sy, ox + 28 * sx, oy + i * 8 * sy)
                c.line(ox + i * 8 * sx, oy, ox + i * 8 * sx, oy + 28 * sy)

        corner_ornament(48, page_h - 76, 1, -1)
        corner_ornament(page_w - 48, page_h - 76, -1, -1)
        corner_ornament(48, 48, 1, 1)
        corner_ornament(page_w - 48, 48, -1, 1)

        def fit_font(text, font, max_size, min_size, max_width):
            for size in range(max_size, min_size - 1, -1):
                if c.stringWidth(text, font, size) <= max_width:
                    return size
            return min_size

        def centred(text, y, font, size, color=ink):
            c.setFillColor(color)
            c.setFont(font, size)
            c.drawCentredString(cx, y, text)

        def divider(y, width=420):
            c.setStrokeColor(gold)
            c.setLineWidth(1)
            c.line(cx - width / 2, y, cx - 18, y)
            c.line(cx + 18, y, cx + width / 2, y)
            c.setFillColor(gold)
            c.circle(cx, y, 3, fill=1, stroke=0)

        # Header
        school_text = (school or "School").strip().upper()
        school_size = fit_font(school_text, "Helvetica-Bold", 20, 12, page_w - 160)
        centred(school_text, page_h - 78, "Helvetica-Bold", school_size, navy)
        divider(page_h - 92, 460)

        centred("CERTIFICATE OF ATTENDANCE", page_h - 128, "Times-Bold", 30, navy)
        centred("OF EXCELLENCE", page_h - 152, "Times-Italic", 13, gold)

        centred("This is to certify that", page_h - 188, "Times-Roman", 14, muted)

        # Recipient name with underline
        name_text = str(name).strip() or "Student"
        name_size = fit_font(name_text, "Times-BoldItalic", 34, 20, page_w - 180)
        centred(name_text, page_h - 228, "Times-BoldItalic", name_size, ink)
        name_w = c.stringWidth(name_text, "Times-BoldItalic", name_size)
        c.setStrokeColor(gold)
        c.setLineWidth(1.2)
        c.line(cx - name_w / 2 - 10, page_h - 238, cx + name_w / 2 + 10, page_h - 238)

        centred(f"Roll No: {roll}    |    Grade: {grade}", page_h - 262, "Helvetica", 11, muted)

        # Body copy
        body_y = page_h - 300
        centred("has demonstrated commendable regularity and commitment by maintaining", body_y, "Times-Roman", 12, ink)
        centred("a consistent attendance record during the academic period.", body_y - 18, "Times-Roman", 12, ink)

        # Stats highlight box
        box_w, box_h = 320, 82
        box_x = cx - box_w / 2
        box_y = page_h - 418
        c.setFillColor(colors.HexColor("#eef4ff"))
        c.setStrokeColor(accent)
        c.setLineWidth(1.5)
        c.roundRect(box_x, box_y, box_w, box_h, 8, fill=1, stroke=1)

        pct_text = f"{pct:.1f}%"
        pct_size = fit_font(pct_text, "Helvetica-Bold", 30, 18, box_w - 40)
        centred(pct_text, box_y + 52, "Helvetica-Bold", pct_size, accent)
        days_label = f"{days} Day{'s' if days != 1 else ''} Present"
        centred("ATTENDANCE ACHIEVED", box_y + 30, "Helvetica-Bold", 9, muted)
        centred(days_label, box_y + 12, "Helvetica", 11, ink)

        centred("This certificate is awarded in recognition of the student's dedication", 168, "Times-Italic", 11, muted)
        centred("to punctual and consistent attendance.", 152, "Times-Italic", 11, muted)

        # Official seal (centre-top of footer band)
        seal_x, seal_y, seal_r = cx, 118, 30
        c.setFillColor(colors.HexColor("#fff8e7"))
        c.setStrokeColor(gold)
        c.setLineWidth(2)
        c.circle(seal_x, seal_y, seal_r, fill=1, stroke=1)
        c.setStrokeColor(navy)
        c.setLineWidth(1)
        c.circle(seal_x, seal_y, seal_r - 5, fill=0, stroke=1)
        c.setFillColor(gold)
        for angle in range(0, 360, 72):
            rad = math.radians(angle - 90)
            c.circle(seal_x + math.cos(rad) * 14, seal_y + math.sin(rad) * 14, 2.5, fill=1, stroke=0)
        c.setFillColor(navy)
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(seal_x, seal_y + 4, "OFFICIAL")
        c.drawCentredString(seal_x, seal_y - 6, "SEAL")

        # Signature blocks
        issue_date = datetime.now().strftime("%d %B %Y")

        def signature_block(x, label):
            c.setStrokeColor(ink)
            c.setLineWidth(0.8)
            c.line(x - 95, 68, x + 95, 68)
            c.setFillColor(ink)
            c.setFont("Helvetica-Bold", 10)
            c.drawCentredString(x, 52, label)
            c.setFillColor(muted)
            c.setFont("Helvetica", 8)
            c.drawCentredString(x, 40, "Authorized Signatory")

        signature_block(cx - 240, "Principal")
        signature_block(cx + 240, "Class Teacher")

        c.setFillColor(muted)
        c.setFont("Helvetica", 9)
        c.drawCentredString(cx - 240, 88, f"Date: {issue_date}")

        c.setFillColor(colors.HexColor("#9ca3af"))
        c.setFont("Helvetica", 7)
        c.drawCentredString(cx, 28, f"Issued via FaceAttend {APP_VERSION}  ·  {issue_date}")

        c.showPage()
        c.save()
        return True

    def _open_exports_folder(self):
        import subprocess, sys as _sys
        EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
        if _sys.platform == "win32":
            subprocess.Popen(["explorer", str(EXPORTS_DIR)])
        else:
            subprocess.Popen(["xdg-open", str(EXPORTS_DIR)])

    def _open_calendar(self):
        """Show attendance calendar heatmap dialog."""
        dlg = tk.Toplevel(self.root)
        dlg.title("📅  " + "calendar")
        dlg.geometry("520x480")
        dlg.configure(bg=P.bg)
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="📅  " + "calendar", bg=P.bg, fg=P.text,
                 font=F_H2).pack(anchor="w", padx=24, pady=(20, 4))
        tk.Label(dlg, text="Visual monthly attendance grid", bg=P.bg, fg=P.muted,
                 font=F_SMALL).pack(anchor="w", padx=24)

        # Month selector
        sel = tk.Frame(dlg, bg=P.bg)
        sel.pack(fill="x", padx=24, pady=(10, 0))
        now = datetime.now()
        month_var = tk.IntVar(value=now.month)
        year_var = tk.IntVar(value=now.year)
        ttk.Spinbox(sel, from_=1, to=12, textvariable=month_var, width=6, font=F_BODY).pack(side="left", padx=(0, 8))
        ttk.Spinbox(sel, from_=2020, to=2030, textvariable=year_var, width=8, font=F_BODY).pack(side="left", padx=(0, 8))

        cal_frame = tk.Frame(dlg, bg=P.bg)
        cal_frame.pack(fill="both", expand=True, padx=24, pady=(10, 16))

        def render():
            for w in cal_frame.winfo_children():
                w.destroy()
            rows = self.store._load_audit_rows()
            days = set()
            for row in rows:
                if row and len(row) > 0:
                    days.add(row[0])
            build_calendar_heatmap(cal_frame, year_var.get(), month_var.get(), days, P, self._lang)

        AnimBtn(sel, text="↺  Render", command=render,
            bg=P.blue, fg="white", activebackground=P.blue_dark,
            activeforeground="white", relief="flat", padx=12, pady=6,
            font=F_SMALL_B, cursor="hand2", anim_color=P.blue_glow).pack(side="left")
        render()

        AnimBtn(dlg, text="✕  Close", command=dlg.destroy,
            bg=P.surface2, fg=P.text, activebackground=P.border,
            activeforeground=P.text, relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2").pack(pady=(0, 14))

    # ── badges / achievements viewer ────────────────────────────────────────
    def _open_badges(self):
        """Show badge and achievement viewer dialog."""
        dlg = tk.Toplevel(self.root)
        dlg.title("🏅  " + "badges")
        dlg.geometry("520x500")
        dlg.configure(bg=P.bg)
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="🏅  " + "badges", bg=P.bg, fg=P.text,
                 font=F_H2).pack(anchor="w", padx=24, pady=(20, 4))
        tk.Label(dlg, text="Student achievements and milestones", bg=P.bg, fg=P.muted,
                 font=F_SMALL).pack(anchor="w", padx=24)

        list_sh = tk.Frame(dlg, bg=P.sidebar2, highlightthickness=1,
                            highlightbackground=P.border_soft)
        list_sh.pack(fill="both", expand=True, padx=24, pady=(10, 10))
        list_sh.grid_rowconfigure(0, weight=1)
        list_sh.grid_columnconfigure(0, weight=1)

        lb = tk.Listbox(list_sh, bg=P.sidebar2, fg=P.text, relief="flat",
                        highlightthickness=0, font=F_BODY, activestyle="none",
                        selectbackground=P.blue_soft, selectforeground=P.text)
        sb = ttk.Scrollbar(list_sh, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        sb.grid(row=0, column=1, sticky="ns")

        students = self.store.load_students()
        for s in students:
            uid = str(s.get("UID", ""))
            name = s.get("Name", "")
            badges = self.badges.get_badges(uid)
            if badges:
                b_str = "  ".join([b["icon"] for b in badges])
                lb.insert("end", f"  {name}  —  {b_str}")
            else:
                lb.insert("end", f"  {name}  —  No badges yet")

        if not students:
            lb.insert("end", "  No students enrolled yet.")

        AnimBtn(dlg, text="✕  Close", command=dlg.destroy,
            bg=P.surface2, fg=P.text, activebackground=P.border,
            activeforeground=P.text, relief="flat", padx=16, pady=9,
            font=F_BODY_B, cursor="hand2").pack(pady=(0, 14))

    # ── whiteboard tab (3.0) ──────────────────────────────────────────────────
    WB_COLORS = ["#0f172a", "#ef4444", "#3b82f6", "#22c55e", "#f59e0b", "#a855f7", "#ffffff"]

    def _build_whiteboard_tab(self):
        side = self._tab_frames["Whiteboard"]
        tool = self._right_tool_frames["Whiteboard"]

        # ── sidebar controls ─────────────────────────────────────────────────
        hdr = tk.Frame(side, bg=P.sidebar)
        hdr.pack(fill="x", padx=18, pady=(18, 10))
        tk.Label(hdr, text="🖍  Whiteboard", bg=P.sidebar, fg=P.text,
                 font=F_H2).pack(anchor="w")
        tk.Label(hdr, text="Freehand notes & diagrams, saved as PNG",
                 bg=P.sidebar, fg=P.muted, font=F_SMALL).pack(anchor="w", pady=(2, 0))

        self._wb_color = tk.StringVar(value=self.WB_COLORS[0])
        self._wb_width = tk.IntVar(value=3)
        self._wb_strokes: list[dict] = []
        self._wb_current: dict | None = None
        self._wb_last_xy = None

        cf = tk.Frame(side, bg=P.sidebar)
        cf.pack(fill="x", padx=18, pady=(4, 4))
        tk.Label(cf, text="Colour", bg=P.sidebar, fg=P.muted, font=F_SMALL_B).pack(anchor="w")
        swatches = tk.Frame(cf, bg=P.sidebar)
        swatches.pack(anchor="w", pady=(4, 0))
        self._wb_swatch_btns = {}
        for c in self.WB_COLORS:
            b = tk.Button(swatches, bg=c, width=2, height=1, relief="flat", bd=0,
                          highlightthickness=2, highlightbackground=P.border,
                          cursor="hand2", command=lambda c=c: self._wb_pick_color(c))
            b.pack(side="left", padx=2)
            self._wb_swatch_btns[c] = b
        AnimBtn(cf, text="🎨 Custom…", command=self._wb_pick_custom_color,
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=10, pady=5,
            font=F_SMALL, cursor="hand2").pack(anchor="w", pady=(8, 0))

        sf = tk.Frame(side, bg=P.sidebar)
        sf.pack(fill="x", padx=18, pady=(12, 4))
        tk.Label(sf, text="Brush size", bg=P.sidebar, fg=P.muted, font=F_SMALL_B).pack(anchor="w")
        tk.Scale(sf, from_=1, to=24, orient="horizontal", variable=self._wb_width,
                 bg=P.sidebar, fg=P.muted, troughcolor=P.surface2, highlightthickness=0,
                 font=F_SMALL, showvalue=True, length=180).pack(anchor="w")

        af = tk.Frame(side, bg=P.sidebar)
        af.pack(fill="x", padx=18, pady=(14, 4))
        AnimBtn(af, text="🧹 Eraser", command=self._wb_use_eraser,
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=10, pady=7,
            font=F_SMALL_B, cursor="hand2").pack(fill="x", pady=(0, 6))
        AnimBtn(af, text="🗑  Clear Board", command=self._wb_clear,
            bg=P.red, fg="white", activebackground=P.red_dark,
            activeforeground="white", relief="flat", padx=10, pady=7,
            font=F_SMALL_B, cursor="hand2", anim_color=P.red_dark).pack(fill="x", pady=(0, 6))
        AnimBtn(af, text="💾  Save as PNG", command=self._wb_save,
            bg=P.green, fg="white", activebackground=P.green_dark,
            activeforeground="white", relief="flat", padx=10, pady=7,
            font=F_SMALL_B, cursor="hand2", anim_color=P.green_dark).pack(fill="x")

        tk.Label(side, text=f"Saved boards go to:\n{WHITEBOARD_DIR}",
                 bg=P.sidebar, fg=P.faint, font=F_SMALL, justify="left",
                 wraplength=260).pack(anchor="w", padx=18, pady=(16, 8))

        # ── full-width canvas ────────────────────────────────────────────────
        thdr = tk.Frame(tool, bg=P.bg)
        thdr.pack(fill="x", padx=24, pady=(18, 8))
        tk.Label(thdr, text="Whiteboard", bg=P.bg, fg=P.text, font=F_H1).pack(anchor="w")
        tk.Label(thdr, text="Click and drag to draw. Pick a colour and size from the panel on the left.",
                 bg=P.bg, fg=P.muted, font=F_BODY).pack(anchor="w")

        canvas_shell = tk.Frame(tool, bg="white", highlightthickness=1,
                                 highlightbackground=P.border)
        canvas_shell.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self._wb_canvas = tk.Canvas(canvas_shell, bg="white", cursor="pencil",
                                     highlightthickness=0)
        self._wb_canvas.pack(fill="both", expand=True)
        self._wb_canvas.bind("<ButtonPress-1>", self._wb_start_stroke)
        self._wb_canvas.bind("<B1-Motion>",     self._wb_draw_stroke)
        self._wb_canvas.bind("<ButtonRelease-1>", self._wb_end_stroke)

    def _wb_pick_color(self, c):
        self._wb_color.set(c)
        for col, b in self._wb_swatch_btns.items():
            b.configure(highlightbackground=P.blue if col == c else P.border)

    def _wb_pick_custom_color(self):
        rgb, hexcode = colorchooser.askcolor(color=self._wb_color.get(), title="Pick a colour")
        if hexcode:
            self._wb_color.set(hexcode)

    def _wb_use_eraser(self):
        # A simple, reliable eraser: draw in the canvas background colour
        # at a larger width, rather than tracking true transparency.
        self._wb_color.set("#ffffff")
        self._wb_width.set(max(self._wb_width.get(), 18))
        Toast.show(self.root, "Eraser selected — draw over what you'd like to remove.", "info")

    def _wb_start_stroke(self, event):
        self._wb_current = {"points": [(event.x, event.y)],
                             "color": self._wb_color.get(),
                             "width": self._wb_width.get()}
        self._wb_strokes.append(self._wb_current)
        self._wb_last_xy = (event.x, event.y)

    def _wb_draw_stroke(self, event):
        if self._wb_current is None or self._wb_last_xy is None:
            return
        x0, y0 = self._wb_last_xy
        self._wb_canvas.create_line(x0, y0, event.x, event.y,
                                     fill=self._wb_current["color"],
                                     width=self._wb_current["width"],
                                     capstyle=tk.ROUND, joinstyle=tk.ROUND, tags="stroke")
        self._wb_current["points"].append((event.x, event.y))
        self._wb_last_xy = (event.x, event.y)

    def _wb_end_stroke(self, _event=None):
        self._wb_current = None
        self._wb_last_xy = None

    def _wb_clear(self):
        if not self._wb_strokes:
            return
        if not messagebox.askyesno("Clear Board", "Clear the whole whiteboard? This can't be undone."):
            return
        self._wb_canvas.delete("stroke")
        self._wb_strokes.clear()
        play_sound(SOUND_NEGATIVE)

    def _wb_save(self):
        if not self._wb_strokes:
            Toast.show(self.root, "Nothing drawn yet.", "warning")
            return
        if Image is None or ImageDraw is None:
            Toast.show(self.root, "Pillow isn't available, so PNG export is disabled.", "error")
            return
        w = max(self._wb_canvas.winfo_width(), 1)
        h = max(self._wb_canvas.winfo_height(), 1)
        img = Image.new("RGB", (w, h), "white")
        draw = ImageDraw.Draw(img)
        for stroke in self._wb_strokes:
            pts = stroke["points"]
            width = stroke["width"]
            color = stroke["color"]
            if len(pts) == 1:
                x, y = pts[0]
                draw.ellipse([x - width / 2, y - width / 2, x + width / 2, y + width / 2], fill=color)
                continue
            for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
                draw.line([x0, y0, x1, y1], fill=color, width=width)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = WHITEBOARD_DIR / f"Whiteboard_{stamp}.png"
        img.save(dest)
        play_sound(SOUND_EXPORT)
        Toast.show(self.root, f"Saved: {dest.name}", "success")
        self.log(f"Whiteboard saved: {dest.name}")

    # ── notepad tab (3.0) ─────────────────────────────────────────────────────
    def _build_notepad_tab(self):
        side = self._tab_frames["Notepad"]
        tool = self._right_tool_frames["Notepad"]

        self._np_current_path: Path | None = None

        hdr = tk.Frame(side, bg=P.sidebar)
        hdr.pack(fill="x", padx=18, pady=(18, 10))
        tk.Label(hdr, text="📝  Notepad", bg=P.sidebar, fg=P.text,
                 font=F_H2).pack(anchor="w")
        tk.Label(hdr, text="Quick notes, saved as .txt files",
                 bg=P.sidebar, fg=P.muted, font=F_SMALL).pack(anchor="w", pady=(2, 0))

        bf = tk.Frame(side, bg=P.sidebar)
        bf.pack(fill="x", padx=18, pady=(4, 8))
        AnimBtn(bf, text="＋ New Note", command=self._np_new,
            bg=P.green, fg="white", activebackground=P.green_dark,
            activeforeground="white", relief="flat", padx=10, pady=7,
            font=F_SMALL_B, cursor="hand2", anim_color=P.green_dark).pack(fill="x", pady=(0, 6))
        AnimBtn(bf, text="🗑 Delete Note", command=self._np_delete,
            bg=P.red, fg="white", activebackground=P.red_dark,
            activeforeground="white", relief="flat", padx=10, pady=7,
            font=F_SMALL_B, cursor="hand2", anim_color=P.red_dark).pack(fill="x")

        tk.Label(side, text="Saved notes", bg=P.sidebar, fg=P.muted,
                 font=F_SMALL_B).pack(anchor="w", padx=18, pady=(10, 2))
        list_sh = tk.Frame(side, bg=P.sidebar2, highlightthickness=1,
                            highlightbackground=P.border)
        list_sh.pack(fill="both", expand=True, padx=18, pady=(0, 12))
        self._np_listbox = tk.Listbox(list_sh, bg=P.sidebar2, fg=P.text,
                                       selectbackground=P.blue, relief="flat",
                                       highlightthickness=0, activestyle="none",
                                       font=F_SMALL)
        self._np_listbox.pack(fill="both", expand=True, padx=6, pady=6)
        self._np_listbox.bind("<Double-1>", lambda _e: self._np_open_selected())

        tk.Label(side, text=f"Saved notes go to:\n{NOTES_DIR}",
                 bg=P.sidebar, fg=P.faint, font=F_SMALL, justify="left",
                 wraplength=260).pack(anchor="w", padx=18, pady=(0, 8))

        # ── full-width editor ────────────────────────────────────────────────
        thdr = tk.Frame(tool, bg=P.bg)
        thdr.pack(fill="x", padx=24, pady=(18, 8))
        self._np_title_lbl = tk.Label(thdr, text="Untitled note", bg=P.bg, fg=P.text, font=F_H1)
        self._np_title_lbl.pack(side="left", anchor="w")
        AnimBtn(thdr, text="💾  Save", command=self._np_save,
            bg=P.green, fg="white", activebackground=P.green_dark,
            activeforeground="white", relief="flat", padx=16, pady=8,
            font=F_BODY_B, cursor="hand2", anim_color=P.green_dark).pack(side="right")

        editor_shell = tk.Frame(tool, bg=P.surface, highlightthickness=1,
                                 highlightbackground=P.border)
        editor_shell.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self._np_text = tk.Text(editor_shell, bg=P.surface, fg=P.text,
                                 insertbackground=P.text, relief="flat",
                                 wrap="word", font=F_BODY, padx=14, pady=12)
        np_sb = ttk.Scrollbar(editor_shell, orient="vertical", command=self._np_text.yview)
        self._np_text.configure(yscrollcommand=np_sb.set)
        self._np_text.pack(side="left", fill="both", expand=True)
        np_sb.pack(side="right", fill="y")

        self._np_refresh_list()

    def _np_refresh_list(self):
        self._np_listbox.delete(0, "end")
        files = sorted(NOTES_DIR.glob("*.txt"), key=lambda p: p.stat().st_mtime, reverse=True)
        self._np_files = files
        for f in files:
            self._np_listbox.insert("end", f.stem)

    def _np_new(self):
        self._np_current_path = None
        self._np_text.delete("1.0", "end")
        self._np_title_lbl.configure(text="Untitled note")

    def _np_open_selected(self):
        sel = self._np_listbox.curselection()
        if not sel:
            return
        path = self._np_files[sel[0]]
        try:
            content = path.read_text(encoding="utf-8")
        except Exception as e:
            Toast.show(self.root, f"Couldn't open note: {e}", "error")
            return
        self._np_current_path = path
        self._np_text.delete("1.0", "end")
        self._np_text.insert("1.0", content)
        self._np_title_lbl.configure(text=path.stem)

    def _np_save(self):
        content = self._np_text.get("1.0", "end-1c")
        if not content.strip():
            Toast.show(self.root, "Nothing to save yet.", "warning")
            return
        if self._np_current_path is None:
            name = simpledialog.askstring("Save Note", "Note name:", parent=self.root)
            if not name or not name.strip():
                return
            safe_name = re.sub(r'[\\/:*?"<>|]+', "_", name.strip())
            self._np_current_path = NOTES_DIR / f"{safe_name}.txt"
        self._np_current_path.write_text(content, encoding="utf-8")
        self._np_title_lbl.configure(text=self._np_current_path.stem)
        self._np_refresh_list()
        play_sound(SOUND_EXPORT)
        Toast.show(self.root, f"Saved: {self._np_current_path.name}", "success")
        self.log(f"Note saved: {self._np_current_path.name}")

    def _np_delete(self):
        sel = self._np_listbox.curselection()
        path = self._np_files[sel[0]] if sel else self._np_current_path
        if path is None or not path.exists():
            Toast.show(self.root, "Select a saved note to delete.", "warning")
            return
        if not messagebox.askyesno("Delete Note", f"Delete '{path.stem}'? This can't be undone."):
            return
        path.unlink()
        if self._np_current_path == path:
            self._np_new()
        self._np_refresh_list()
        play_sound(SOUND_NEGATIVE)
        Toast.show(self.root, "Note deleted.", "info")

    # ── console tab ───────────────────────────────────────────────────────────
    def _build_console_tab(self):
        parent = self._tab_frames["Console"]
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        hdr = tk.Frame(parent, bg=P.sidebar)
        hdr.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        tk.Label(hdr, text="📋  Console", bg=P.sidebar, fg=P.text,
                 font=F_H2).pack(side="left")
        AnimBtn(hdr, text="Clear", command=self._console_clear,
            bg=P.red, fg="white", activebackground=P.red_dark,
            activeforeground="white", relief="flat", padx=10, pady=4,
            font=F_SMALL_B, cursor="hand2",
            anim_color="#b91c1c").pack(side="right")

        lsh = tk.Frame(parent, bg="#070d16", highlightthickness=1, highlightbackground=P.border_soft)
        lsh.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        lsh.grid_rowconfigure(0, weight=1)
        lsh.grid_columnconfigure(0, weight=1)
        self.console_box = tk.Text(lsh, wrap="word", bg="#070d16", fg="#c8d8ec",
                                   relief="flat", highlightthickness=0,
                                   font=("Consolas", 9), padx=10, pady=10, state="disabled")
        cs = ttk.Scrollbar(lsh, orient="vertical", command=self.console_box.yview)
        self.console_box.configure(yscrollcommand=cs.set)
        self.console_box.grid(row=0, column=0, sticky="nsew")
        cs.grid(row=0, column=1, sticky="ns")

        cf = tk.Frame(parent, bg=P.sidebar)
        cf.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 16))
        cf.grid_columnconfigure(0, weight=1)
        self._cmd_var = tk.StringVar()
        ce = ttk.Entry(cf, textvariable=self._cmd_var, font=("Consolas", 10))
        ce.grid(row=0, column=0, sticky="ew")
        ce.bind("<Return>", lambda _: self._run_cmd())
        AnimBtn(cf, text="Run", command=self._run_cmd,
            bg=P.blue, fg="white", activebackground=P.blue_dark,
            activeforeground="white", relief="flat", padx=12, pady=7,
            font=("Segoe UI", 9, "bold"), cursor="hand2",
            anim_color=P.blue_glow).grid(row=0, column=1, padx=(8, 0))

    def _clog(self, text):
        self.console_box.configure(state="normal")
        self.console_box.insert("end", text + "\n")
        self.console_box.see("end")
        self.console_box.configure(state="disabled")

    def _console_clear(self):
        self.console_box.configure(state="normal")
        self.console_box.delete("1.0", "end")
        self.console_box.configure(state="disabled")

    def _run_cmd(self):
        raw = self._cmd_var.get().strip()
        self._cmd_var.set("")
        if not raw: return
        self._clog(f"» {raw}")
        cmd = raw.lower()

        if cmd in ("/list", "list"):
            students = self.store.load_students()
            self._clog(f"  {len(students)} student(s):")
            for s in students:
                self._clog(f"    {s.get('UID','')} | {s.get('Name','')} | {s.get('Grade','')} | {s.get('No. of Days Present',0)} days")

        elif cmd in ("/add", "add"):
            self.open_add_student()

        elif cmd in ("/refresh", "refresh"):
            self.refresh_all(); self._clog("  Refreshed.")

        elif cmd in ("/scan start", "scan start"):
            self.start_scan()

        elif cmd in ("/scan stop", "scan stop"):
            self.stop_scan()

        elif cmd in ("/stats", "stats"):
            students = self.store.load_students()
            present  = len([u for u in self.scanned_today if u in self.students_by_uid])
            top = max(students, key=lambda s: safe_int(s.get("XP")), default=None)
            self._clog(f"  Total: {len(students)} | Present: {present} | Absent: {len(students)-present}")
            if top:
                self._clog(f"  XP leader: {top.get('Name','')} ({top.get('XP',0)} XP)")

        elif cmd in ("/backup", "backup"):
            p = self.store.backup()
            self._clog(f"  Backup saved: {p.name}")
            Toast.show(self.root, f"Backup saved: {p.name}", "success")

        elif cmd.startswith("/export"):
            self._do_export("excel")

        elif cmd.startswith("/remove ") or cmd.startswith("remove "):
            q = re.sub(r"^/?remove\s+(@s\s+)?", "", raw, flags=re.I).strip()
            matches = self._find_students(q)
            if not matches: self._clog(f"  No student matching '{q}'.")
            elif len(matches) > 1:
                self._clog(f"  Multiple matches for '{q}':")
                for m in matches: self._clog(f"    {m.get('UID','')} | {m.get('Name','')}")
            else:
                s   = matches[0]; uid = str(s.get("UID","")); name = str(s.get("Name",uid))
                if messagebox.askyesno("Remove", f"Remove {name} ({uid})?"):
                    self.store.remove_student(uid)
                    self.scanned_today.discard(uid)
                    self._save_today_scans()
                    self.refresh_all()
                    self._clog(f"  Removed {name}.")

        elif cmd.startswith("/edit ") or cmd.startswith("edit "):
            q = re.sub(r"^/?edit\s+(@s\s+)?", "", raw, flags=re.I).strip()
            matches = self._find_students(q)
            if not matches: self._clog(f"  No student matching '{q}'.")
            elif len(matches) > 1: self._clog(f"  Multiple matches. Be more specific.")
            else:
                self.open_edit_student_by_uid(str(matches[0].get("UID", "")))

        elif cmd.startswith("/msg ") or cmd.startswith("msg "):
            q = re.sub(r"^/?msg\s+(@s\s+)?", "", raw, flags=re.I).strip()
            matches = self._find_students(q)
            if not matches: self._clog(f"  No student matching '{q}'.")
            elif len(matches) > 1: self._clog("  Multiple matches. Be more specific.")
            else:
                s = matches[0]
                lbl = f"{s.get('Name','')}  ({s.get('Roll No','')} / {s.get('UID','')})"
                self._msg_refresh_list()
                self._msg_stu_var.set(lbl)
                self._msg_to_email.set(str(s.get("Parent Email", "")))
                self._msg_to_phone.set(str(s.get("Parent Phone", "")))
                self._switch("Mail"); self._clog(f"  Opened Messages for {s.get('Name','')}.")

        elif cmd in ("/theme", "theme"):
            self._theme_dark = not self._theme_dark
            set_theme(self._theme_dark, self._lang)
            self.settings.set("theme_dark", self._theme_dark)
            self.settings.save()
            self._clog(f"  Theme: {'Dark' if self._theme_dark else 'Light'}")
            Toast.show(self.root, f"Theme: {'Dark' if self._theme_dark else 'Light'} (restart to apply fully)", "info")

        elif cmd in ("/api", "api"):
            if self.api_server._running:
                self.api_server.stop()
                self._clog(f"  API stopped.")
            else:
                if self.api_server.start():
                    self._clog(f"  API started on port {API_PORT}")
                else:
                    self._clog("  API failed to start.")

        elif cmd in ("/notify-absent", "notify-absent", "/notifyabsent"):
            absentees = self.get_absent_students()
            if not absentees:
                self._clog("  No absent students right now (or school day inactive).")
            else:
                result = self.send_absent_notifications(manual=True)
                play_sound(SOUND_EXPORT)
                self._clog(f"  Absentee notify: {result['sent_email']} email(s), "
                           f" {result['skipped']} skipped, "
                           f"{result['total']} absent total.")
                Toast.show(self.root, f"Notified parents of {result['total']} absent student(s).", "success")

        elif cmd in ("/restart", "restart"):
            self._clog("  Restarting FaceAttend…")
            self._perform_restart()

        elif cmd in ("/badge", "badge"):
            students = self.store.load_students()
            self._clog("  Badge standings:")
            for s in students[:10]:
                uid = str(s.get("UID", ""))
                badges = self.badges.get_badges(uid)
                b_str = ", ".join([b["icon"] + " " + _(b["name"], self._lang) for b in badges]) or "None"
                self._clog(f"    {s.get('Name','')} ({uid}): {b_str}")

        elif cmd in ("/?", "/help", "help"):
            cmds = [
                "/list          — list all students",
                "/add           — open Add Student",
                "/edit @s <n>   — edit a student",
                "/remove @s <n> — remove a student",
                "/msg @s <n>    — open Messages for a student",
                "/refresh       — refresh dashboard",
                "/scan start    — start camera scan",
                "/scan stop     — stop camera scan",
                "/stats         — show quick stats",
                "/backup        — backup Excel file",
                "/export        — export styled Excel",
                "/notify-absent — send absence notice to all absent parents now",
                "/theme         — toggle dark/light theme",
                "/restart       — restart FaceAttend now (Windows)",
                "/api           — toggle REST API server",
                "/badge         — show badge standings",
                "/clear         — clear this console",
                "/?             — show this help",
            ]
            self._clog("  ┌─ Commands ──────────────────────────────")
            for c in cmds: self._clog(f"  │  {c}")
            self._clog("  └────────────────────────────────────────")

        elif cmd in ("/clear", "clear"):
            self._console_clear()

        else:
            self._clog("  Unknown command. Type /? for help.")

    # ── settings tab ─────────────────────────────────────────────────────────
    def _build_settings_tab(self):
        parent = self._tab_frames["Settings"]
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # Make scrollable
        outer = tk.Frame(parent, bg=P.sidebar)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_rowconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)
        canvas = tk.Canvas(outer, bg=P.sidebar, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        shell = tk.Frame(canvas, bg=P.sidebar)
        win_id = canvas.create_window((0, 0), window=shell, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win_id, width=e.width))
        shell.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        hdr = tk.Frame(shell, bg=P.sidebar)
        hdr.pack(fill="x", padx=18, pady=(18, 10))
        tk.Label(hdr, text="⚙  " + "Settings", bg=P.sidebar, fg=P.text,
                 font=F_H2).pack(anchor="w")
        tk.Label(hdr, text="Customise FaceAttend behaviour", bg=P.sidebar,
                 fg=P.muted, font=F_SMALL).pack(anchor="w", pady=(2, 0))

        def card():
            f = tk.Frame(shell, bg=P.sidebar2, highlightthickness=1, highlightbackground=P.border)
            f.pack(fill="x", padx=18, pady=(0, 4))
            return f

        def section(title, subtitle):
            f = tk.Frame(shell, bg=P.sidebar)
            f.pack(fill="x", padx=18, pady=(10, 4))
            tk.Label(f, text=title, bg=P.sidebar, fg=P.text, font=F_H3).pack(anchor="w")
            if subtitle:
                tk.Label(f, text=subtitle, bg=P.sidebar, fg=P.faint, font=F_TINY).pack(anchor="w")

        def row(c, txt, r, var, width=None, show=None):
            tk.Label(c, text=txt, bg=P.sidebar2, fg=P.muted,
                     font=F_SMALL_B).grid(row=r, column=0, sticky="w", padx=14, pady=(10,4))
            kw = {"textvariable": var, "font": F_BODY}
            if width: kw["width"] = width
            if show: kw["show"] = show
            ttk.Entry(c, **kw).grid(row=r, column=1, sticky="e" if width else "ew", padx=14, pady=(10,4))

        # ── Appearance ────────────────────────────────────────────
        section("Appearance", "Theme & language — changing either will offer an automatic restart")
        c_appearance = card()
        c_appearance.grid_columnconfigure(1, weight=1)
        self._theme_var = tk.BooleanVar(value=self._theme_dark)
        tk.Label(c_appearance, text="Theme", bg=P.sidebar2, fg=P.muted,
                 font=F_SMALL_B).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 10))
        theme_cb = ttk.Combobox(c_appearance, values=["Dark", "Light"],
                                 state="readonly", font=F_BODY)
        theme_cb.set("Dark" if self._theme_dark else "Light")
        theme_cb.grid(row=0, column=1, sticky="ew", padx=14, pady=(10, 10))
        theme_cb.bind("<<ComboboxSelected>>", self._on_theme_change)

        tk.Label(c_appearance, text="Language", bg=P.sidebar2, fg=P.muted,
                 font=F_SMALL_B).grid(row=1, column=0, sticky="w", padx=14, pady=(0, 10))
        self._lang_code_by_label = {v: k for k, v in LANG_NAMES.items()}
        lang_labels = list(LANG_NAMES.values())
        lang_cb = ttk.Combobox(c_appearance, values=lang_labels,
                                state="readonly", font=F_BODY)
        lang_cb.set(LANG_NAMES.get(self._lang, "English"))
        lang_cb.grid(row=1, column=1, sticky="ew", padx=14, pady=(0, 10))
        lang_cb.bind("<<ComboboxSelected>>", self._on_language_change)
        self._lang_cb = lang_cb

        self._settings_restart_btn = AnimBtn(c_appearance, text="⟳  Restart Now to Apply",
            command=self._perform_restart,
            bg=P.blue, fg="white", activebackground=P.blue_dark,
            activeforeground="white", relief="flat", padx=14, pady=8,
            font=F_SMALL_B, cursor="hand2", anim_color=P.blue_glow)
        self._settings_restart_btn.grid(row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(2, 12))
        self._settings_restart_btn.grid_remove()  # hidden until a theme/language change happens

        # ── Attendance Rules ────────────────────────────────────
        section("Attendance Rules", "Late cutoff, recognition sensitivity & XP rewards")
        c1 = card()
        self._late_var      = tk.StringVar(value=self.settings.late_time)
        self._threshold_var = tk.StringVar(value=str(self.settings.threshold))
        self._xp_on_var     = tk.StringVar(value=str(self.settings.xp_ontime))
        self._xp_late_var   = tk.StringVar(value=str(self.settings.xp_late))
        self._school_var    = tk.StringVar(value=self.settings.school_name or "")
        row(c1, "Late cutoff (HH:MM)", 0, self._late_var,      width=8)
        row(c1, "School name",         1, self._school_var)
        row(c1, "Face threshold",      2, self._threshold_var,  width=8)
        row(c1, "XP (on-time)",        3, self._xp_on_var,      width=8)
        row(c1, "XP (late)",           4, self._xp_late_var,    width=8)

        # ── Schedule Manager ────────────────────────────────────
        section("Schedule", "Per-day late thresholds and active/inactive days")
        c_sched = card()
        self._sched_vars = {}
        days = ScheduleManager.DAYS
        for i, day in enumerate(days):
            d = self.schedule._data.get(day, {})
            active_var = tk.BooleanVar(value=d.get("active", day not in ("Sat", "Sun")))
            late_var = tk.StringVar(value=d.get("late_time", "08:32"))
            self._sched_vars[day] = (active_var, late_var)
            fr = tk.Frame(c_sched, bg=P.sidebar2)
            fr.grid(row=i, column=0, sticky="ew", padx=14, pady=2)
            fr.grid_columnconfigure(1, weight=1)
            ttk.Checkbutton(fr, text=day, variable=active_var).grid(row=0, column=0, sticky="w")
            ttk.Entry(fr, textvariable=late_var, width=8, font=F_BODY).grid(row=0, column=1, sticky="w", padx=(10, 0))

        # ── Camera & Scanning ───────────────────────────────────
        section("Camera & Scanning", "Multi-camera and detection options")
        c_cam = card()
        self._cam_var = tk.StringVar(value=str(self.settings.camera_index))
        self._liveness_var = tk.BooleanVar(value=self.settings.liveness_check)
        tk.Label(c_cam, text="Camera", bg=P.sidebar2, fg=P.muted,
                 font=F_SMALL_B).grid(row=0, column=0, sticky="w", padx=14, pady=(10,4))
        cam_opts = [f"Camera {i}" for i in self._cameras]
        cam_cb = ttk.Combobox(c_cam, values=cam_opts, state="readonly", font=F_BODY,
                               textvariable=self._cam_var)
        cam_cb.grid(row=0, column=1, sticky="ew", padx=14, pady=(10,4))
        ttk.Checkbutton(c_cam, text="Liveness Detection",
                         variable=self._liveness_var).grid(
            row=1, column=0, columnspan=2, sticky="w", padx=14, pady=(4, 4))

        # ── Email Sender ────────────────────────────────────────
        section("Email Sender", "Used by the Messages tab to send parent emails via Gmail")
        c2 = card()
        self._smtp_sender_var = tk.StringVar(value=self.settings.smtp_sender or "")
        self._smtp_pw_var     = tk.StringVar(value=self.settings.smtp_password or "")
        row(c2, "Gmail address", 0, self._smtp_sender_var)
        row(c2, "App password", 1, self._smtp_pw_var, show="•")
        tk.Label(c2, text="Generate at myaccount.google.com/apppasswords",
                 bg=P.sidebar2, fg=P.faint,
                 font=F_TINY).grid(row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(0,10))

        # ── Auto-Notify Parents ─────────────────────────────────
        section("Auto-Notify Parents", "Send a message automatically the moment a student is scanned")
        c3 = card()
        self._auto_email_var    = tk.BooleanVar(value=bool(self.settings.auto_notify_email))

        def _toggle_email(*_):
            self.settings.set("auto_notify_email", bool(self._auto_email_var.get()))
            self.settings.save()
        self._auto_email_var.trace_add("write", _toggle_email)

        ttk.Checkbutton(c3, text="Email parent automatically on scan",
                         variable=self._auto_email_var).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(12, 4))
        tk.Label(c3, text="Fires once per student per day, right after attendance is marked.\n"
                          "Email needs the Gmail sender above. WhatsApp opens/reuses a\n"
                          "web.whatsapp.com tab \u2014 keep it logged in for this to work.",
                 bg=P.sidebar2, fg=P.faint, font=F_TINY, justify="left").grid(
            row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(2, 10))

        # ── Absentee Alerts ───────────────────────────────────────
        section("Absentee Alerts", "Notify parents of students who haven't been marked present")
        c_abs = card()
        self._absent_enabled_var  = tk.BooleanVar(value=bool(self.settings.absent_notify_enabled))
        self._absent_email_var    = tk.BooleanVar(value=bool(self.settings.absent_notify_email))
        self._absent_time_var     = tk.StringVar(value=str(self.settings.absent_notify_time or "09:30"))

        def _toggle_abs_enabled(*_):
            self.settings.set("absent_notify_enabled", bool(self._absent_enabled_var.get()))
            self.settings.save()
        def _toggle_abs_email(*_):
            self.settings.set("absent_notify_email", bool(self._absent_email_var.get()))
            self.settings.save()
        self._absent_enabled_var.trace_add("write", _toggle_abs_enabled)
        self._absent_email_var.trace_add("write", _toggle_abs_email)

        ttk.Checkbutton(c_abs, text="Automatically send absentee alerts at a fixed time every school day",
                         variable=self._absent_enabled_var).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(12, 4))

        tk.Label(c_abs, text="Fixed time (HH:MM, 24h)", bg=P.sidebar2, fg=P.muted,
                 font=F_SMALL_B).grid(row=1, column=0, sticky="w", padx=14, pady=(2, 4))
        time_entry = ttk.Entry(c_abs, textvariable=self._absent_time_var, width=8, font=F_BODY)
        time_entry.grid(row=1, column=1, sticky="w", pady=(2, 4))

        def _save_abs_time(*_):
            val = self._absent_time_var.get().strip()
            try:
                h, m = val.split(":")
                int(h); int(m)
                self.settings.set("absent_notify_time", f"{int(h):02d}:{int(m):02d}")
                self.settings.save()
            except Exception:
                pass
        time_entry.bind("<FocusOut>", _save_abs_time)
        time_entry.bind("<Return>", _save_abs_time)

        ttk.Checkbutton(c_abs, text="via Email",
                         variable=self._absent_email_var).grid(
            row=2, column=0, sticky="w", padx=14, pady=(4, 4))

        AnimBtn(c_abs, text="📨  Send to All Absent Parents Now", command=self._send_absent_now,
            bg=P.amber, fg="#1a1100", activebackground="#d97706",
            activeforeground="#1a1100", relief="flat", padx=14, pady=8,
            font=F_SMALL_B, cursor="hand2", anim_color="#b45309").grid(
                row=3, column=0, columnspan=3, sticky="w", padx=14, pady=(8, 4))

        tk.Label(c_abs, text="Checks who is NOT yet marked present today and sends each parent\n"
                            "the absence message below (edit in the message template if needed).\n"
                            "Runs once per day automatically if enabled, or any time via this button\n"
                            "or the /notify-absent console command.",
                 bg=P.sidebar2, fg=P.faint, font=F_TINY, justify="left").grid(
            row=4, column=0, columnspan=3, sticky="w", padx=14, pady=(2, 10))

        # ── REST API ────────────────────────────────────────────
        section("REST API", "External integration on port 8765")
        c_api = card()
        self._api_var = tk.BooleanVar(value=self.settings.api_server)
        ttk.Checkbutton(c_api, text="Enable REST API server",
                         variable=self._api_var).grid(
            row=0, column=0, sticky="w", padx=14, pady=(12, 4))
        self._api_status_lbl = tk.Label(c_api, text="Stopped", bg=P.sidebar2, fg=P.red,
                                         font=F_SMALL_B)
        self._api_status_lbl.grid(row=0, column=1, sticky="e", padx=14, pady=(12, 4))
        tk.Label(c_api, text="Endpoints: /api/students, /api/today, /api/houses, /api/qr?uid=...",
                 bg=P.sidebar2, fg=P.faint, font=F_TINY).grid(
            row=1, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 10))

        # ── Data & Backup ───────────────────────────────────────
        section("Data & Backup", "Save your work before trying something risky")
        bf = tk.Frame(shell, bg=P.sidebar)
        bf.pack(fill="x", padx=18, pady=(0, 10))
        AnimBtn(bf, text="💾  Save Settings", command=self._save_settings,
            bg=P.green, fg="white", activebackground=P.green_dark,
            activeforeground="white", relief="flat", padx=14, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color="#15803d").pack(side="left", padx=(0, 8))
        AnimBtn(bf, text="↺  Reset", command=self._reset_settings,
            bg=P.sidebar2, fg=P.ink, activebackground=P.border,
            activeforeground=P.ink, relief="flat", padx=14, pady=9,
            font=F_BODY_B, cursor="hand2").pack(side="left", padx=(0, 8))
        AnimBtn(bf, text="📦  Backup", command=lambda: (
            Toast.show(self.root, f"Backup: {self.store.backup().name}", "success"),
            self.log("Excel backed up.")
        ), bg=P.blue, fg="white", activebackground=P.blue_dark,
            activeforeground="white", relief="flat", padx=14, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color=P.blue_glow).pack(side="left")

        # ── Tools ───────────────────────────────────────────────
        section("Tools", "Bulk import, certificates, and emergency tools")
        tf = tk.Frame(shell, bg=P.sidebar)
        tf.pack(fill="x", padx=18, pady=(0, 18))
        AnimBtn(tf, text="📥  Bulk Import", command=self._open_bulk_import,
            bg=P.blue, fg="white", activebackground=P.blue_dark,
            activeforeground="white", relief="flat", padx=14, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color=P.blue_glow).pack(side="left", padx=(0, 8))
        AnimBtn(tf, text="🏅  Certificates", command=self._open_certificates,
            bg=P.amber, fg="#1a1100", activebackground="#d97706",
            activeforeground="#1a1100", relief="flat", padx=14, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color="#b45309").pack(side="left", padx=(0, 8))
        AnimBtn(tf, text="📅  Calendar", command=self._open_calendar,
            bg=P.purple, fg="white", activebackground=P.purple,
            activeforeground="white", relief="flat", padx=14, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color="#7c3aed").pack(side="left")

        section("Year-End Promotion", "Move every student up a grade in one go — "
                                       "run this once a year instead of editing students one by one")
        pf = tk.Frame(shell, bg=P.sidebar)
        pf.pack(fill="x", padx=18, pady=(0, 18))
        AnimBtn(pf, text="🎓  Promote Students (New Year)", command=self._open_promotion_dialog,
            bg=P.green_dark, fg="white", activebackground="#15803d",
            activeforeground="white", relief="flat", padx=14, pady=9,
            font=F_BODY_B, cursor="hand2", anim_color="#0f6e2c").pack(side="left")
        tk.Label(pf, text="e.g. 7C → 8C for the whole school, with per-section overrides",
                 bg=P.sidebar, fg=P.faint, font=F_SMALL).pack(side="left", padx=(10, 0))

    def _open_promotion_dialog(self):
        PromotionDialog(self)

    def _save_settings(self):
        global LATE_TIME, FACE_THRESHOLD
        try:
            h, m = self._late_var.get().strip().split(":")
            LATE_TIME = dtime(int(h), int(m))
            self.settings.set("late_time", f"{int(h):02d}:{int(m):02d}")
        except Exception:
            Toast.show(self.root, "Invalid time format. Use HH:MM", "error")
            play_sound(SOUND_ERROR); return
        try:
            FACE_THRESHOLD = float(self._threshold_var.get())
            self.settings.set("threshold", FACE_THRESHOLD)
        except Exception:
            Toast.show(self.root, "Invalid threshold. Use e.g. 0.50", "error")
            play_sound(SOUND_ERROR); return
        self.settings.set("xp_ontime",     safe_int(self._xp_on_var.get(),  10))
        self.settings.set("xp_late",       safe_int(self._xp_late_var.get(), 2))
        self.settings.set("school_name",   self._school_var.get().strip())
        self.settings.set("smtp_sender",   self._smtp_sender_var.get().strip())
        self.settings.set("smtp_password", self._smtp_pw_var.get().strip())
        self.settings.set("auto_notify_email",    bool(self._auto_email_var.get()))
        # 2.4 new settings
        self.settings.set("liveness_check", bool(self._liveness_var.get()))
        try:
            cam_idx = int(self._cam_var.get().replace("Camera ", ""))
            self.settings.set("camera_index", cam_idx)
        except Exception:
            pass
        self.settings.set("api_server", bool(self._api_var.get()))
        # Schedule
        for day, (active_var, late_var) in self._sched_vars.items():
            self.schedule.set_day(day, bool(active_var.get()), late_var.get().strip())
        self.settings.save()
        if hasattr(self, "_msg_from_var"):
            self._msg_from_var.set(self.settings.smtp_sender or "(set in Settings)")
        # API toggle
        if self.settings.api_server and not self.api_server._running:
            self.api_server.start()
            self._api_status_lbl.configure(text="Running", fg=P.green)
            self.log(f"API server started on port {API_PORT}")
        elif not self.settings.api_server and self.api_server._running:
            self.api_server.stop()
            self._api_status_lbl.configure(text="Stopped", fg=P.red)
            self.log("API server stopped")
        Toast.show(self.root, "settings_saved", "success")
        self.log("Settings saved.")

    def _reset_settings(self):
        global LATE_TIME, FACE_THRESHOLD
        LATE_TIME = dtime(8, 32); FACE_THRESHOLD = 0.50
        self._late_var.set("08:32"); self._threshold_var.set("0.50")
        self._xp_on_var.set("10"); self._xp_late_var.set("2")
        self._school_var.set(""); self._smtp_sender_var.set(""); self._smtp_pw_var.set("")
        self._auto_email_var.set(False)
        self._liveness_var.set(True)
        self._api_var.set(False)
        if hasattr(self, "_absent_enabled_var"):  self._absent_enabled_var.set(False)
        if hasattr(self, "_absent_email_var"):    self._absent_email_var.set(True)
        if hasattr(self, "_absent_whatsapp_var"): self._absent_whatsapp_var.set(False)
        if hasattr(self, "_absent_time_var"):     self._absent_time_var.set("09:30")
        self.settings._data = dict(self.settings._defaults)
        self.settings.save()
        self._theme_dark = self.settings.theme_dark
        self._lang = "en"
        set_theme(self._theme_dark, self._lang)
        if hasattr(self, "_lang_cb"):
            self._lang_cb.set(LANG_NAMES.get("en", "English"))
        if hasattr(self, "_msg_from_var"):
            self._msg_from_var.set("(set in Settings)")
        Toast.show(self.root, "settings_reset", "info")
        self.log("Settings reset to defaults.")
        self._offer_restart("Settings reset to defaults.")

    def _on_theme_change(self, event=None):
        dark = event.widget.get() == "Dark"
        self._theme_dark = dark
        self.settings.set("theme_dark", dark)
        set_theme(dark, self._lang)
        self.settings.save()
        self._offer_restart("Theme changed.")

    def _on_language_change(self, event=None):
        label = self._lang_cb.get()
        code  = self._lang_code_by_label.get(label, "en")
        self._lang = code
        self.settings.set("language", code)
        self.settings.save()
        play_sound(SOUND_SUCCESS)
        self._offer_restart(f"Language set to {label}.")
        self.log(f"Language changed to {label} ({code}).")

    def _offer_restart(self, reason: str):
        """
        Shows a toast with a Restart button after a Theme/Language change.
        Clicking it (or the matching button in Settings) rebuilds the app
        in place — works on every platform now, not just Windows.
        """
        Toast.show(self.root, f"{reason} Click Restart to apply it now.", "info", 8000,
                   action_label="Restart", action_command=self._perform_restart)
        if hasattr(self, "_settings_restart_btn"):
            self._settings_restart_btn.grid()

    def _perform_restart(self):
        """
        Rebuilds the whole app UI in place, in the same process and the
        same window, instead of spawning a separate detached process.

        3.0 fix: the previous implementation spawned a brand-new FaceAttend
        process (Windows-only) and then destroyed this window, trusting
        that the new process would appear before/as the old one vanished.
        In practice that hand-off could fail silently (blocked process
        creation, antivirus, a race over the locked Excel file while the
        old process was still shutting down) — from the teacher's side
        that just looks like "the app closed". Rebuilding in place can't
        fail that way: there is only ever one process and one window, so
        a Theme/Language change is guaranteed to actually show up rather
        than gamble on a second process launching correctly.
        """
        overlay = tk.Toplevel(self.root)
        overlay.overrideredirect(True)
        overlay.attributes("-topmost", True)
        overlay.configure(bg=P.sidebar)
        w, h = 320, 130
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        overlay.geometry(f"{w}x{h}+{(sw - w)//2}+{(sh - h)//2}")
        shell = tk.Frame(overlay, bg=P.sidebar, highlightthickness=1, highlightbackground=P.border_soft)
        shell.pack(fill="both", expand=True, padx=2, pady=2)
        tk.Label(shell, text="⟳", bg=P.sidebar, fg=P.blue, font=("Segoe UI", 26)).pack(pady=(18, 4))
        tk.Label(shell, text="Restarting FaceAttend…", bg=P.sidebar, fg=P.text,
                 font=F_BODY_B).pack()
        overlay.update()

        # Stop anything running in the background under the OLD instance
        # before we tear its widgets down, so a stray timer/thread doesn't
        # fire against destroyed widgets.
        self._restarting = True
        try:
            if self.scanning:
                self.stop_scan()
        except Exception as e:
            self.log(f"Restart: error stopping scan: {e}")
        try:
            self.api_server.stop()
        except Exception:
            pass

        self.root.after(200, lambda: self._rebuild_in_place(overlay))

    def _rebuild_in_place(self, overlay):
        try:
            overlay.destroy()
        except tk.TclError:
            pass
        for w in list(self.root.winfo_children()):
            try:
                w.destroy()
            except Exception:
                pass
        try:
            FaceAttendApp(self.root)
        except Exception as e:
            logger.error("Error rebuilding FaceAttend after restart: %s", e)
            messagebox.showerror("FaceAttend",
                f"FaceAttend hit an error while restarting:\n\n{e}")

    # ── camera ────────────────────────────────────────────────────────────────
    def toggle_scan(self):
        if self.scanning: self.stop_scan()
        else:             self.start_scan()

    def start_scan(self):
        if np is None:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "NumPy missing. Install requirements.", "error"); return
        if not ensure_face_recognition():
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "face_recognition not available.", "error"); return
        cam_idx = self.settings.camera_index
        self.camera = open_camera(cam_idx)
        if self.camera is None or not self.camera.isOpened():
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Cannot open webcam.", "error"); return
        self.camera.set(cv2.CAP_PROP_FRAME_WIDTH,  CAMERA_W)
        self.camera.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_H)
        self.scanning = True
        self.scan_btn.configure(text=_("stop_scan", self._lang), bg=P.red, activebackground=P.red_dark)
        self._set_scanning_indicator(True)
        play_sound(SOUND_CAMERA_START)
        self.log("Scan started.")
        self.liveness.reset()
        self._camera_loop()

    def stop_scan(self):
        self.scanning = False
        self.scan_btn.configure(text=_("start_scan", self._lang), bg=P.blue, activebackground=P.blue_dark)
        self._set_scanning_indicator(False)
        if self.camera:
            self.camera.release(); self.camera = None
        self.cam_lbl.configure(image="", text="📷  Camera is off")
        self.log("Scan stopped.")

    def _camera_loop(self):
        if not self.scanning or self.camera is None:
            return
        try:
            ok, frame = self.camera.read()
            if ok:
                self.last_frame = frame.copy()
                self.frame_count += 1
                # ── QR scan mode ──────────────────────────────────────
                if self._qr_mode:
                    qr_data = decode_qr_from_frame(frame)
                    if qr_data and "uid" in qr_data:
                        uid = str(qr_data["uid"])
                        self._handle_recognition(uid)
                        self._qr_mode = False
                # ── Face processing ───────────────────────────────────
                if self.frame_count % FRAME_SKIP == 0:
                    try:
                        self.last_results = self._process(frame)
                    except Exception as e:
                        self.last_results = []
                        self.log(f"Frame process error: {e}")
                # ── Liveness check overlay ────────────────────────────
                try:
                    ann = self._draw(frame.copy(), self.last_results)
                    if self._liveness_enabled and self.liveness._predictor is not None:
                        live = self.liveness.is_live(frame)
                        color = (34, 197, 94) if live else (239, 68, 68)
                        txt = "LIVE" if live else "SPOOF?"
                        cv2.putText(ann, txt, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)
                except Exception:
                    ann = frame.copy()
                self._show_frame(ann)
        except Exception as e:
            self.log(f"Camera loop error: {e}")
        self.root.after(20, self._camera_loop)

    def _process(self, frame) -> list[dict]:
        import time as _t
        small    = cv2.resize(frame, (0,0), fx=0.25, fy=0.25)
        rgb      = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
        locs     = face_recognition.face_locations(rgb, model="hog")
        encs     = face_recognition.face_encodings(rgb, locs)
        results  = []
        for loc, enc in zip(locs, encs):
            t, r, b, l = [v * 4 for v in loc]
            uid = name = None; dist = None; status = "unknown"
            if self.store.known_encodings:
                try:
                    dists = face_recognition.face_distance(self.store.known_encodings, enc)
                    bi    = int(np.argmin(dists)); bd = float(dists[bi])
                    dist  = bd
                    if bd <= FACE_THRESHOLD:
                        uid    = self.store.known_uids[bi]
                        name   = str(self.students_by_uid.get(uid, {}).get("Name", uid))
                        status = "known"
                        self._handle_recognition(uid)
                except Exception as e:
                    import traceback
                    self.log(f"Recognition error: {e}")
                    print("".join(traceback.format_exception(type(e), e, e.__traceback__)))

            if status == "unknown":
                now = _t.time()
                if now - self._unknown_last_saved >= 10.0:
                    self._unknown_last_saved = now
                    self._save_unknown_face(frame, t, r, b, l)

            results.append({"box": (t,r,b,l), "uid": uid,
                             "name": name or "Unknown", "distance": dist, "status": status})
        return results

    def _save_unknown_face(self, frame, t, r, b, l):
        """Crop and save an unrecognised face to data/unknowns/ with a timestamp filename."""
        try:
            UNKNOWNS_DIR.mkdir(parents=True, exist_ok=True)
            pad  = 20
            h, w = frame.shape[:2]
            crop = frame[max(0, t-pad):min(h, b+pad), max(0, l-pad):min(w, r+pad)]
            if crop.size == 0:
                return
            # Resize if oversized to keep unknowns folder tidy
            if crop.shape[0] > 400 or crop.shape[1] > 400:
                scale = 400 / max(crop.shape[0], crop.shape[1])
                crop = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            path  = UNKNOWNS_DIR / f"unknown_{stamp}.jpg"
            cv2.imwrite(str(path), crop)
            self.log(f"⚠ Unknown face saved → {path.name}")
        except Exception as e:
            self.log(f"Unknown face save failed: {e}")

    def _draw(self, frame, results):
        for r in results:
            t, ri, b, l = r["box"]
            known = r["status"] == "known"
            color = (34, 197, 94) if known else (0, 175, 255)
            lbl   = r["name"]
            if r["distance"] is not None and known:
                lbl = f"{lbl}  {r['distance']:.2f}"
            cv2.rectangle(frame, (l,t), (ri,b), color, 2)
            cv2.rectangle(frame, (l, max(0,t-28)), (ri,t), color, cv2.FILLED)
            cv2.putText(frame, lbl, (l+6, max(18,t-8)),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.58, (255,255,255), 1, cv2.LINE_AA)
        return frame

    def _show_frame(self, frame):
        if Image is None or cv2 is None:
            return
        try:
            rgb  = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img  = Image.fromarray(rgb)
            img.thumbnail((424, 310))
            canvas = Image.new("RGB", (424, 310), "#070d16")
            canvas.paste(img, ((424-img.width)//2, (310-img.height)//2))
            self.camera_photo = ImageTk.PhotoImage(canvas)
            self.cam_lbl.configure(image=self.camera_photo, text="")
        except Exception as e:
            self.log(f"Camera display error: {e}")

    def get_capture_frame(self):
        if self.last_frame is not None: return self.last_frame.copy()
        tmp = open_camera()
        if tmp is None or not tmp.isOpened(): return None
        ok, f = tmp.read(); tmp.release()
        return f if ok else None

    def _handle_recognition(self, uid: str):
        now = datetime.now()
        if uid in self.scanned_today:
            last = self.last_dup_log.get(uid)
            if last is None or (now - last).total_seconds() > 12:
                name = self.students_by_uid.get(uid, {}).get("Name", uid)
                self.log(f"{name} already marked today.")
                self.last_dup_log[uid] = now
            return
        # ── Schedule check ──────────────────────────────────────
        if not self.schedule.is_active():
            self.log("Today is not an active school day.")
            Toast.show(self.root, "Today is not an active school day.", "warning")
            return
        is_late = now.time() > self.schedule.late_time()
        try:
            student, xp = self.store.record_attendance(uid, is_late, self.settings)
        except ValueError as e:
            self.log(str(e)); return
        self.scanned_today.add(uid)
        self._save_today_scans()
        self.students_by_uid = self.store.students_by_uid()
        name   = str(student.get("Name", uid))
        roll   = str(student.get("Roll No", ""))
        status = "attendance_late" if is_late else "attendance_ontime"
        # ── Voice feedback removed in v2.9 ──────────────────────
        # ── Badge awards ────────────────────────────────────────
        new_badges = self.badges.check_and_award(uid, student, is_early=not is_late)
        for badge in new_badges:
            badge_name = _(badge["name"], self._lang)
            Toast.show(self.root, f"🏅 New badge: {badge['icon']} {badge_name}!", "success", 5000)
            self.log(f"Badge awarded to {name}: {badge_name}")
        self.last_name.configure(text=f"Welcome, {name}")
        self.last_detail.configure(text=f"{status}  ·  {roll}  ·  +{xp} XP  ·  {now.strftime('%I:%M %p')}")
        self._animate_result_card()
        Toast.show(self.root, f"{name} marked {status.lower()}  (+{xp} XP)", "success" if not is_late else "warning")
        self.log(f"{name} marked {status.lower()}, +{xp} XP")
        if self.settings.auto_notify_email or self.settings.auto_notify_whatsapp:
            try:
                self._auto_notify_parent(student, status, xp, now)
            except Exception as e:
                self.log(f"Auto-notify error: {e}")
        self.refresh_all()

    def _animate_result_card(self):
        """Flash the result card border blue briefly."""
        seq  = [P.blue, P.green, P.green, P.border]
        def step(i):
            if i < len(seq):
                try:
                    self.result_card.configure(highlightbackground=seq[i])
                    self.root.after(120, lambda: step(i+1))
                except tk.TclError:
                    pass
        step(0)

    # ── student actions ───────────────────────────────────────────────────────
    def open_add_student(self):
        play_sound(SOUND_SUCCESS)
        StudentDialog(self)

    def open_edit_student(self):
        sel = self.table.selection()
        if not sel:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Select a student row first.", "warning"); return
        if len(sel) > 1:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Edit works on one student at a time — "
                                   "use Remove or Change Section for bulk actions.", "warning")
            return
        vals = self.table.item(sel[0], "values")
        uid  = str(vals[2])
        self.open_edit_student_by_uid(uid)

    # ── 3.0: multi-select helpers ────────────────────────────────────────────
    def _selected_uids(self) -> list[str]:
        """UID + display-name pairs for every currently selected roster row."""
        out = []
        for row_id in self.table.selection():
            vals = self.table.item(row_id, "values")
            if len(vals) > 3:
                out.append((str(vals[2]), str(vals[3])))
        return out

    def _on_table_selection_changed(self, _event=None):
        n = len(self.table.selection())
        if hasattr(self, "selection_count_lbl"):
            self.selection_count_lbl.configure(
                text=f"{n} selected" if n > 1 else "")

    def select_all_students(self):
        """Select every row currently shown in the table (respects an
        active search filter — only visible/filtered rows are selected)."""
        rows = self.table.get_children("")
        if not rows:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "No students to select.", "warning")
            return
        self.table.selection_set(rows)
        Toast.show(self.root, f"Selected all {len(rows)} student(s).", "info")

    def open_bulk_section_change(self):
        """
        Change the Grade/Section field (e.g. '7C') for every currently
        selected student in one go — the same underlying action the
        Promotion tool uses, but for an ad-hoc manual reassignment
        (e.g. a couple of students moving between sections mid-year).
        """
        pairs = self._selected_uids()
        if not pairs:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Select one or more students first.", "warning")
            return
        new_section = simpledialog.askstring(
            "Change Section",
            f"New Grade/Section for {len(pairs)} selected student(s) "
            f"(e.g. 8C):", parent=self.root)
        if not new_section or not new_section.strip():
            return
        new_section = new_section.strip().upper()
        uids = [uid for uid, _name in pairs]
        self.store.bulk_set_grades({uid: new_section for uid in uids})
        self.students_by_uid = self.store.students_by_uid()
        self.refresh_all()
        play_sound(SOUND_SUCCESS)
        Toast.show(self.root, f"Moved {len(uids)} student(s) to {new_section}.", "success")
        self.log(f"Bulk section change: {len(uids)} student(s) → {new_section}.")

    def open_edit_student_by_uid(self, uid: str):
        s = self.students_by_uid.get(uid)
        if s is None:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Student not found.", "error"); return
        StudentDialog(self, student=s)

    def mark_attendance_selected(self):
        """
        Manual attendance fallback for when face-scan fails (e.g. camera
        issue, or a student's face won't recognise). Teacher selects the
        row in the table, clicks this button, and the student is marked
        present right now — same on-time/late XP rule as a real face scan,
        and the resulting record looks identical to a scanned one (no
        'manual' tag), per spec.
        """
        sel = self.table.selection()
        if not sel:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Select a student row first.", "warning"); return
        vals = self.table.item(sel[0], "values")
        uid, name = str(vals[2]), str(vals[3])

        if uid in self.scanned_today:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, f"{name} is already marked present today.", "warning")
            return
        if not self.schedule.is_active():
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Today is not an active school day.", "warning")
            return

        now = datetime.now()
        is_late = now.time() > self.schedule.late_time()
        try:
            student, xp = self.store.record_attendance(uid, is_late, self.settings)
        except ValueError as e:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, str(e), "error")
            return

        self.scanned_today.add(uid)
        self._save_today_scans()
        self.students_by_uid = self.store.students_by_uid()
        roll   = str(student.get("Roll No", ""))
        status = "attendance_late" if is_late else "attendance_ontime"

        new_badges = self.badges.check_and_award(uid, student, is_early=not is_late)
        for badge in new_badges:
            badge_name = _(badge["name"], self._lang)
            Toast.show(self.root, f"🏅 New badge: {badge['icon']} {badge_name}!", "success", 5000)
            self.log(f"Badge awarded to {name}: {badge_name}")

        self.last_name.configure(text=f"Welcome, {name}")
        self.last_detail.configure(text=f"{status}  ·  {roll}  ·  +{xp} XP  ·  {now.strftime('%I:%M %p')}  ·  manually marked")
        self._animate_result_card()
        play_sound(SOUND_SUCCESS)
        Toast.show(self.root, f"{name} manually marked {status.lower()}  (+{xp} XP)",
                   "success" if not is_late else "warning")
        self.log(f"{name} manually marked {status.lower()}, +{xp} XP")

        if self.settings.auto_notify_email or self.settings.auto_notify_whatsapp:
            try:
                self._auto_notify_parent(student, status, xp, now)
            except Exception as e:
                self.log(f"Auto-notify error: {e}")
        self.refresh_all()

    def _send_absent_now(self):
        absentees = self.get_absent_students()
        if not absentees:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "No absent students right now (or today isn't an active school day).", "warning")
            return
        if not self.settings.absent_notify_email and not self.settings.absent_notify_whatsapp:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Enable Email and/or WhatsApp for absentee alerts first.", "warning")
            return
        if not messagebox.askyesno("Send Absentee Alerts",
                                    f"Send an absence notice to the parents of {len(absentees)} "
                                    f"student(s) not yet marked present today?"):
            return
        result = self.send_absent_notifications(manual=True)
        play_sound(SOUND_EXPORT)
        Toast.show(self.root,
                   f"Notified parents of {result['total']} absent student(s) "
                   f"({result['sent_email']} email).",
                   "success", 5500)

    def remove_selected(self):
        pairs = self._selected_uids()
        if not pairs:
            play_sound(SOUND_ERROR)
            Toast.show(self.root, "Select a student to remove.", "warning"); return
        play_sound(SOUND_NEGATIVE)
        if len(pairs) == 1:
            uid, name = pairs[0]
            prompt = f"Remove {name} ({uid})?"
        else:
            names = ", ".join(n for _u, n in pairs[:8])
            more  = "" if len(pairs) <= 8 else f", and {len(pairs) - 8} more"
            prompt = f"Remove {len(pairs)} students — {names}{more}?"
        if not messagebox.askyesno("Remove Student" if len(pairs) == 1 else "Remove Students", prompt):
            return
        uids = [uid for uid, _name in pairs]
        self.store.bulk_remove(uids)
        for uid in uids:
            self.scanned_today.discard(uid)
        self._save_today_scans()
        self.refresh_all()
        if len(pairs) == 1:
            Toast.show(self.root, f"{pairs[0][1]} removed.", "info")
            self.log(f"Removed {pairs[0][1]} ({pairs[0][0]}).")
        else:
            Toast.show(self.root, f"Removed {len(pairs)} students.", "info")
            self.log(f"Bulk removed {len(pairs)} students: " + ", ".join(uids) + ".")

    # ── export ────────────────────────────────────────────────────────────────
    def _export_menu(self):
        if not hasattr(self, "_export_menu_ref"):
            self._export_menu_ref = tk.Menu(self.root, tearoff=False, bg=P.surface, fg=P.text,
                                            activebackground=P.blue, activeforeground="white",
                                            relief="flat", font=("Segoe UI", 10))
        m = self._export_menu_ref
        m.delete(0, "end")
        m.add_command(label="📊  Export to Excel (styled)",
                      command=lambda: self._do_export("excel"))
        m.add_command(label="📄  Export to CSV",
                      command=lambda: self._do_export("csv"))
        try:
            m.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            m.grab_release()

    def _do_export(self, fmt: str):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if fmt == "excel":
            dest = EXPORTS_DIR / f"FaceAttend_export_{stamp}.xlsx"
            p    = self.store.export_excel(dest)
            if p:
                play_sound(SOUND_EXPORT)
                Toast.show(self.root, f"Excel exported: {p.name}", "success")
                self.log(f"Exported Excel: {p.name}")
            else:
                play_sound(SOUND_ERROR)
                Toast.show(self.root, "Export failed.", "error")
        else:
            dest = EXPORTS_DIR / f"FaceAttend_export_{stamp}.csv"
            p    = self.store.export_csv(dest)
            play_sound(SOUND_EXPORT)
            Toast.show(self.root, f"CSV exported: {p.name}", "success")
            self.log(f"Exported CSV: {p.name}")


    # ── refresh ───────────────────────────────────────────────────────────────
    def refresh_all(self):
        self.students_by_uid = self.store.students_by_uid()
        self._refresh_stats()
        self.refresh_table(self.search_var.get().strip())
        self._msg_refresh_list()
        if hasattr(self, "_an_top_list"):
            self._refresh_analytics()
        if hasattr(self, "_house_cards"):
            self._refresh_houses()

    def _refresh_stats(self):
        students = list(self.students_by_uid.values())
        total    = len(students)
        present  = len([u for u in self.scanned_today if u in self.students_by_uid])
        absent   = max(0, total - present)
        top      = max(students, key=lambda s: safe_int(s.get("XP")), default=None)
        self.card_total.set(total)
        self.card_present.set(present)
        self.card_absent.set(absent)
        self.card_xp.set(f"{top.get('XP',0)} XP" if top else "—")
        self.card_xp.set_subtitle(str(top.get("Name", "")) if top else "No scans yet")
        self.date_lbl.configure(text=datetime.now().strftime("%A, %d %B %Y"))

    def refresh_table(self, query=""):
        ql = query.lower()
        self.table.delete(*self.table.get_children())
        self.table_photos = []
        self._hovered_row = None
        students = self.store.load_students()

        # apply column sort if active
        if self._sort_col and self._sort_col in COLUMNS:
            def sort_key(s):
                v = s.get(self._sort_col, "")
                try: return (0, float(v))
                except (TypeError, ValueError): return (1, str(v).lower())
            students = sorted(students, key=sort_key, reverse=not self._sort_ascending)

        # update heading arrows
        for col in COLUMNS:
            arrow = ""
            if col == self._sort_col:
                arrow = "  ↑" if self._sort_ascending else "  ↓"
            self.table.heading(col, text=col + arrow)
        shown = 0

        def _flame(streak: int) -> str:
            s = safe_int(streak)
            if s >= 30: return "🔥🔥🔥"
            if s >= 10: return "🔥🔥"
            if s >= 5:  return "🔥"
            return ""

        for i, s in enumerate(students):
            uid    = str(s.get("UID", ""))
            streak = safe_int(s.get("Streak", 0))
            flame  = _flame(streak)
            name   = s.get("Name", "")
            name_display = f"{flame} {name}".strip() if flame else name
            house  = str(s.get("House","None")).strip()
            vals = [
                s.get("Sl.No",""), s.get("Roll No",""), uid, name_display,
                s.get("Grade",""), house,
                s.get("Parent Phone",""), s.get("Parent Email",""),
                s.get("No. of Days Present",0), s.get("XP",0),
                f"{safe_float(s.get('Multiplier'),1.0):.1f}",
                streak,
            ]
            hay  = " ".join(str(v) for v in vals).lower()
            if ql and ql not in hay: continue
            tags = ["present"] if uid in self.scanned_today else ["even" if i%2==0 else "odd"]
            if ql and ql in hay and uid not in self.scanned_today: tags = ["match"]
            photo = self._thumbnail(uid)
            if photo: self.table_photos.append(photo)
            self.table.insert("", "end", text="", image=photo or "",
                              values=vals, tags=tags)
            shown += 1
        self.table.update_idletasks()

        if hasattr(self, "search_count_lbl"):
            self.search_count_lbl.configure(text=f"{shown} found" if ql else "")

        if hasattr(self, "table_empty"):
            if not students:
                self.table_empty.place(relx=0.5, rely=0.42, anchor="center")
            else:
                self.table_empty.place_forget()

    def _thumbnail(self, uid):
        if Image is None: return None
        fp = FACES_DIR / f"{uid}.jpg"
        if not fp.exists(): return None
        try:
            img = Image.open(fp); img.thumbnail((60, 60))
            canvas = Image.new("RGB", (62, 62), P.surface2)
            canvas.paste(img.convert("RGB"), ((62-img.width)//2, (62-img.height)//2))
            return ImageTk.PhotoImage(canvas)
        except: return None

    def _debounce_search(self, _e=None):
        if self._search_after:
            self.root.after_cancel(self._search_after)
        self._search_after = self.root.after(300, lambda: self.refresh_table(self.search_var.get().strip()))

    def _clear_search(self):
        self.search_var.set("")
        self.refresh_table("")

    # ── today scans persistence ───────────────────────────────────────────────
    def _load_today_scans(self) -> set[str]:
        if not TODAY_SCANS_FILE.exists(): return set()
        try:
            with open(TODAY_SCANS_FILE, "rb") as f: p = pickle.load(f)
            if p.get("date") != date.today().isoformat(): return set()
            return set(p.get("uids", []))
        except: return set()

    def _save_today_scans(self):
        with open(TODAY_SCANS_FILE, "wb") as f:
            pickle.dump({"date": date.today().isoformat(), "uids": sorted(self.scanned_today)}, f)

    # ── logging ───────────────────────────────────────────────────────────────
    def log(self, text: str):
        stamp = datetime.now().strftime("%H:%M:%S")
        if self.activity.size() == 1 and self.activity.get(0) == self._feed_empty_hint:
            self.activity.delete(0)
        self.activity.insert(0, f"{stamp}  {text}")
        while self.activity.size() > 150:
            self.activity.delete(tk.END)
        if hasattr(self, "console_box"):
            self._clog(f"{stamp}  {text}")

    def _on_close(self):
        try:
            if self.scanning:
                self.stop_scan()
        except Exception as e:
            self.log(f"Close error: {e}")
        finally:
            self.root.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  SPLASH  (animated dot-progress)
# ═══════════════════════════════════════════════════════════════════════════════
class SplashScreen:
    """
    Professional loading screen that shows each import/dependency check
    as a live row (✓ green = ok, ⚠ amber = missing) instead of terminal
    typewriter output. Completes in ~5-7 seconds.

    3.0 fix: the window used to be a fixed 420px tall regardless of how
    many checklist rows there were, so on machines where the row font
    rendered a little taller (or a future update adds more checks) rows
    quietly overflowed past the bottom edge of the window — nothing was
    actually scrolling, they were just invisible, which looked like the
    splash "scrolled down and did nothing" and occasionally left the
    window stuck since the visible content stopped updating. The window
    now sizes itself to the actual number of rows, and a watchdog timer
    guarantees the app reaches the main window even if a single check
    throws or a step callback never fires.
    """

    # Each entry: (display_label, check_fn_or_None)
    # check_fn returns True = ok, False = missing/warning
    CHECKS = [
        ("Core Python modules",        lambda: True),
        ("Tkinter UI",                 lambda: True),
        ("OpenCV (cv2)",               lambda: cv2 is not None),
        ("NumPy",                      lambda: np is not None),
        ("Pillow (PIL)",               lambda: Image is not None),
        ("openpyxl (Excel)",           lambda: Workbook is not None),
        ("face_recognition",           lambda: face_recognition is not None),
        ("Matplotlib charts",          lambda: plt is not None),
        ("pygame (sound)",             lambda: _check_pygame()),
        ("pyttsx3 (voice)",            lambda: pyttsx3 is not None),
        ("reportlab (PDF)",            lambda: HAS_REPORTLAB),
        ("Excel database",             lambda: EXCEL_FILE.exists() or True),
        ("Face encodings",             lambda: ENCODINGS_FILE.exists() or True),
        (f"FaceAttend {APP_VERSION}",  lambda: True),
    ]

    ROW_DELAY     = 380     # ms between each row appearing
    FADE_STEPS    = 10      # steps for fade-in
    ROW_HEIGHT_PX = 22      # used only to size the window, not to lay out rows
    WATCHDOG_MS   = 12000   # hard ceiling — always reach the main window by then

    def __init__(self, master: tk.Toplevel, on_complete):
        self.master      = master
        self.on_complete = on_complete
        self._step        = 0
        self._finished     = False
        self._row_labels  = []   # (icon_lbl, text_lbl) per row
        play_sound(SOUND_STARTUP)
        self._build_ui()
        self.master.after(50, self._fade_in)
        # Safety net: no matter what happens to the row-by-row chain
        # above, guarantee we reach the main window within WATCHDOG_MS.
        self.master.after(self.WATCHDOG_MS, self._watchdog_fire)

    def _build_ui(self):
        self.master.configure(bg="#0e1117")
        # Size the window to fit every row instead of a fixed height that
        # rows could silently overflow past — this is the core startup fix.
        header_h   = 96
        progress_h = 72
        margins    = 40
        content_h  = len(self.CHECKS) * self.ROW_HEIGHT_PX
        W = 680
        H = min(header_h + progress_h + margins + content_h, 760)
        H = max(H, 380)
        sw = self.master.winfo_screenwidth()
        sh = self.master.winfo_screenheight()
        self.master.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")
        self.master.attributes("-alpha", 0.0)

        # Outer shell with subtle border
        shell = tk.Frame(self.master, bg="#12161f",
                         highlightthickness=1, highlightbackground="#2a3448")
        shell.pack(fill="both", expand=True, padx=2, pady=2)

        # ── header ──────────────────────────────────────────────
        hdr = tk.Frame(shell, bg="#12161f")
        hdr.pack(fill="x", padx=28, pady=(24, 0))

        tk.Label(hdr, text="⬡", bg="#12161f", fg="#3b82f6",
                 font=("Segoe UI", 26)).pack(side="left", padx=(0, 12))

        title_wrap = tk.Frame(hdr, bg="#12161f")
        title_wrap.pack(side="left")
        tk.Label(title_wrap, text=f"FaceAttend  {APP_VERSION}",
                 bg="#12161f", fg="#f1f5f9",
                 font=("Segoe UI", 20, "bold")).pack(anchor="w")
        tk.Label(title_wrap, text=f"Build {APP_BUILD}  ·  Smart Attendance System",
                 bg="#12161f", fg="#7b8aa0",
                 font=("Segoe UI", 9)).pack(anchor="w")

        # thin accent line under header
        tk.Frame(shell, bg="#2a3448", height=1).pack(fill="x", padx=28, pady=(14, 0))

        # ── checklist area (every row is guaranteed to fit — see H above) ────
        self._list_frame = tk.Frame(shell, bg="#12161f")
        self._list_frame.pack(fill="both", expand=True, padx=28, pady=(10, 0))

        # pre-create all rows hidden, reveal one-by-one
        for label, _ in self.CHECKS:
            row = tk.Frame(self._list_frame, bg="#12161f")
            row.pack(fill="x", pady=1)
            row.pack_forget()   # hidden until revealed

            icon_lbl = tk.Label(row, text="·", bg="#12161f", fg="#4d586c",
                                 font=("Segoe UI", 11), width=2)
            icon_lbl.pack(side="left", padx=(0, 6))
            txt_lbl = tk.Label(row, text=label, bg="#12161f", fg="#4d586c",
                                font=("Segoe UI", 10))
            txt_lbl.pack(side="left")

            self._row_labels.append((row, icon_lbl, txt_lbl))

        # ── progress bar ──────────────────────────────────────────
        bar_frame = tk.Frame(shell, bg="#12161f")
        bar_frame.pack(fill="x", padx=28, pady=(10, 14))

        style = ttk.Style(self.master)
        style.configure("Splash.Horizontal.TProgressbar",
            troughcolor="#1c2233", background="#3b82f6",
            bordercolor="#1c2233", lightcolor="#3b82f6",
            darkcolor="#2563eb", thickness=5)
        self.pb = ttk.Progressbar(bar_frame,
                                   style="Splash.Horizontal.TProgressbar",
                                   orient="horizontal", mode="determinate",
                                   maximum=len(self.CHECKS), value=0)
        self.pb.pack(fill="x")

        self.status_var = tk.StringVar(value="Starting up…")
        tk.Label(bar_frame, textvariable=self.status_var,
                 bg="#12161f", fg="#4d586c",
                 font=("Segoe UI", 8)).pack(anchor="e", pady=(4, 0))

    def _fade_in(self, alpha=0.0):
        alpha += 0.1
        try:
            self.master.attributes("-alpha", min(1.0, alpha))
        except tk.TclError:
            return
        if alpha < 1.0:
            self.master.after(25, lambda: self._fade_in(alpha))
        else:
            self.master.after(150, self._advance)

    def _advance(self):
        if self._finished:
            return
        if self._step >= len(self.CHECKS):
            self.master.after(300, self._done)
            return

        row, icon_lbl, txt_lbl = self._row_labels[self._step]
        label, check_fn = self.CHECKS[self._step]

        # show the row
        row.pack(fill="x", pady=1)

        # run check and colour accordingly — a failing/throwing check no
        # longer stalls the sequence, it's just shown as a warning row
        try:
            ok = check_fn() if check_fn else True
        except Exception:
            ok = False

        if ok:
            icon_lbl.configure(text="✓", fg="#22c55e")
            txt_lbl.configure(fg="#e8edf5")
        else:
            icon_lbl.configure(text="⚠", fg="#f59e0b")
            txt_lbl.configure(fg="#f59e0b")

        self.pb["value"] = self._step + 1
        self.status_var.set(label)
        self._step += 1
        self.master.after(self.ROW_DELAY, self._advance)

    def _watchdog_fire(self):
        """If _advance()'s after() chain never reached _done() for any
        reason, force completion so the app never appears to hang on the
        splash screen indefinitely."""
        if not self._finished:
            self._done()

    def _done(self):
        if self._finished:
            return
        self.status_var.set("Ready.")
        self.master.after(300, self._close)

    def _close(self):
        if self._finished:
            return
        self._finished = True
        try:
            self.master.destroy()
        except tk.TclError:
            pass
        try:
            self.on_complete()
        except Exception as e:
            logger.error("Error while launching main window: %s", e)
            try:
                messagebox.showerror("FaceAttend",
                    f"FaceAttend hit an error while starting up:\n\n{e}\n\n"
                    f"Check face_recognition_error.log next to the app for details.")
            except Exception:
                pass


def _check_pygame() -> bool:
    try:
        import pygame
        return True
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY
# ═══════════════════════════════════════════════════════════════════════════════
def _prepare():
    ensure_dirs()
    if not EXCEL_FILE.exists() and Workbook:
        AttendanceStore(Settings())._init_excel()

def main():
    _prepare()
    root = tk.Tk()
    root.withdraw()
    root.title(f"FaceAttend {APP_VERSION}")
    root.geometry("1440x860")
    root.minsize(1100, 720)
    root.configure(bg=P.bg)
    if ICON_FILE.exists():
        try:
            root.iconbitmap(default=str(ICON_FILE))
        except Exception:
            pass

    def launch():
        try:
            root.deiconify()
            if os.name == "nt":
                root.state("zoomed")
            FaceAttendApp(root)
        except Exception as e:
            logger.error("Failed to launch main window: %s", e)
            try:
                messagebox.showerror("FaceAttend",
                    f"FaceAttend couldn't start:\n\n{e}")
            except Exception:
                pass
            root.deiconify()

    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.resizable(False, False)
    splash.attributes("-topmost", True)
    # SplashScreen destroys `splash` itself once its checklist/watchdog
    # finishes, then calls on_complete() — don't destroy it again here.
    SplashScreen(splash, launch)
    root.mainloop()

if __name__ == "__main__":
    main()
