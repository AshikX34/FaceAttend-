#!/usr/bin/env python3
# ══════════════════════════════════════════════════════════════════════════
#  FaceAttend 3.2 — web edition
#
#  A local web app. The interface runs in Firefox; the camera, face
#  recognition and Excel storage all run here in Python.
#
#  Reads and writes the SAME files as FaceAttend 3.x:
#      FaceAttend_Students.xlsx, data/faces/, data/encodings.pkl
#  so your existing roster carries straight over.
#
#  Run:  python3 faceattend_web.py
# ══════════════════════════════════════════════════════════════════════════
from __future__ import annotations

import base64
import csv
import io
import json
import os
import pickle
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import date, datetime, timedelta
from pathlib import Path

APP_VERSION = "4.0"
APP_BUILD = "2026.08-crossplatform"

# ══════════════════════════════════════════════════════════════════════════
#  FIRST-RUN BOOTSTRAP
#  Installs what's missing, then relaunches itself. Makes this file the only
#  thing anyone needs — no batch file, no setup script, no README steps.
#  Skip with:  --no-install
# ══════════════════════════════════════════════════════════════════════════
CORE_PACKAGES = {          # import name -> pip name
    "flask": "flask",
    "cv2": "opencv-python",
    "numpy": "numpy",
    "PIL": "Pillow",
    "openpyxl": "openpyxl",
}
# Gives the native desktop window. If it can't be installed the app still
# runs perfectly - it just opens in a browser instead.
WINDOW_PACKAGES = {
    "webview": "pywebview",
}
# On Windows and macOS the OS supplies the webview itself (WebView2 /
# WKWebView), so pywebview is small and always worth having — it becomes a
# core requirement there rather than an optional extra.
if sys.platform.startswith("win") or sys.platform == "darwin":
    CORE_PACKAGES["webview"] = "pywebview"
FACE_PACKAGES = {
    "face_recognition": "face_recognition",
}


def _missing_imports(mapping: dict) -> list[str]:
    import importlib.util
    out = []
    for mod, pipname in mapping.items():
        if importlib.util.find_spec(mod) is None:
            out.append(pipname)
    return out


def _bootstrap() -> None:
    """Install core packages on first run, then restart into a working app."""
    if "--no-install" in sys.argv or os.getenv("FACEATTEND_NO_INSTALL"):
        return
    if os.getenv("FACEATTEND_BOOTSTRAPPED"):
        return                      # already tried once; don't loop

    need = _missing_imports(CORE_PACKAGES)
    want_window = _missing_imports(WINDOW_PACKAGES) if "--browser" not in sys.argv else []
    # face_recognition pulls in dlib, which compiles from C++ source and can
    # take 10-25 minutes. Silently blocking first launch on that is worse than
    # starting without it — the app runs fine and marks attendance manually.
    # Opt in explicitly with --setup-face.
    want_face = _missing_imports(FACE_PACKAGES) if "--setup-face" in sys.argv else []
    if not need and not want_window and not want_face:
        return
    if not need and not want_face and want_window:
        # Only the window package is missing - try quietly, don't block startup.
        import subprocess as _sp
        try:
            _sp.call([sys.executable, "-m", "pip", "install", "--quiet", *want_window],
                     stdout=_sp.DEVNULL, stderr=_sp.DEVNULL, timeout=180)
        except Exception:
            pass
        return

    if want_face:
        # face_recognition (dlib) is the fragile one, especially on Windows:
        # newer setuptools breaks pkg_resources, which dlib's build needs.
        # Pin setuptools first so the dlib wheel/sdist build doesn't explode.
        import subprocess as _sp
        print("  Installing face recognition engine (this can take a few minutes)...")
        try:
            _sp.call([sys.executable, "-m", "pip", "install", "--quiet",
                      "--force-reinstall", "setuptools==69.5.1"])
        except Exception:
            pass

    need = need + want_window + want_face

    print()
    print("  First run - FaceAttend needs a few Python packages:")
    for n in need:
        print(f"      {n}")
    print()
    print("  Installing now. This usually takes 1-3 minutes.")
    print("  (skip this with:  python faceattend_web.py --no-install)")
    print()

    import subprocess as _sp
    cmd = [sys.executable, "-m", "pip", "install", "--quiet", *need]
    # Debian/Ubuntu system Pythons refuse installs without this flag (PEP 668).
    in_venv = hasattr(sys, "real_prefix") or sys.base_prefix != sys.prefix
    try:
        rc = _sp.call(cmd)
        if rc != 0 and not in_venv:
            print("  Retrying with --break-system-packages ...")
            rc = _sp.call(cmd + ["--break-system-packages"])
    except Exception as exc:
        print(f"  Could not run pip: {exc}")
        rc = 1

    if rc != 0:
        print()
        print("  Automatic install failed. Please run this yourself:")
        print(f"      {sys.executable} -m pip install " + " ".join(need))
        print()
        if sys.platform.startswith("win"):
            try:
                input("  Press Enter to close...")
            except Exception:
                pass
        sys.exit(1)

    print("  Packages installed - restarting FaceAttend...")
    print()
    env = dict(os.environ, FACEATTEND_BOOTSTRAPPED="1")
    os.execve(sys.executable, [sys.executable, os.path.abspath(__file__), *sys.argv[1:]], env)


_bootstrap()

# ── optional heavy deps, all degrade gracefully ──────────────────────────
MISSING: list[str] = []
try:
    import cv2
except Exception:
    cv2 = None; MISSING.append("opencv-python")
try:
    import numpy as np
except Exception:
    np = None; MISSING.append("numpy")
try:
    import face_recognition
except Exception:
    face_recognition = None; MISSING.append("face_recognition")
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except Exception:
    Workbook = load_workbook = None; MISSING.append("openpyxl")
try:
    from flask import Flask, Response, jsonify, request, send_file, send_from_directory
except Exception:
    print("FATAL: Flask is required.  pip install flask")
    raise

# ── paths ────────────────────────────────────────────────────────────────
IS_WINDOWS = sys.platform.startswith("win")
IS_MAC     = sys.platform == "darwin"


def _resolve_home() -> Path:
    """
    Where FaceAttend keeps its data.

    Beside the script when that folder is writable (the normal case on
    Windows and for a downloaded script), otherwise a per-user location:
      Windows  %LOCALAPPDATA%/FaceAttend
      macOS    ~/Library/Application Support/FaceAttend
      Linux    ~/.local/share/faceattend
    """
    env = os.getenv("FACEATTEND_HOME")
    if env:
        return Path(env).expanduser().resolve()

    here = Path(__file__).resolve().parent
    system_prefixes = ("/usr", "/opt", "/srv", "/snap") if not IS_WINDOWS else ()
    looks_system = bool(system_prefixes) and str(here).startswith(system_prefixes)
    if not looks_system:
        try:
            if os.access(here, os.W_OK):
                return here
        except Exception:
            pass

    if IS_WINDOWS:
        base = os.getenv("LOCALAPPDATA") or os.getenv("APPDATA") or str(Path.home())
        return Path(base) / "FaceAttend"
    if IS_MAC:
        return Path.home() / "Library" / "Application Support" / "FaceAttend"
    xdg = os.getenv("XDG_DATA_HOME") or str(Path.home() / ".local" / "share")
    return Path(xdg).expanduser().resolve() / "faceattend"

APP_DIR       = _resolve_home()
DATA_DIR      = APP_DIR / "data"
FACES_DIR     = DATA_DIR / "faces"
UNKNOWNS_DIR  = DATA_DIR / "unknowns"
LOGS_DIR      = DATA_DIR / "logs"
BACKUPS_DIR   = APP_DIR / "backups"
EXPORTS_DIR   = APP_DIR / "exports"
BOARD_DIR     = APP_DIR / "Whiteboard"
NOTES_DIR     = APP_DIR / "Notes"
SETTINGS_FILE = DATA_DIR / "settings.json"
SCANS_FILE    = DATA_DIR / "today_scans.pkl"
ENCODINGS_FILE = DATA_DIR / "encodings.pkl"
AUDIT_LOG     = LOGS_DIR / "attendance_log.csv"
EXCEL_FILE    = APP_DIR / "FaceAttend_Students.xlsx"
BADGE_FILE    = DATA_DIR / "badges.json"
CERTS_DIR     = APP_DIR / "Certificates"
QR_DIR        = APP_DIR / "QRCodes"
SLIPS_DIR     = APP_DIR / "Slips"
EVIDENCE_DIR  = DATA_DIR / "evidence"
QUEUE_FILE    = DATA_DIR / "pending_scans.json"
KEY_FILE      = DATA_DIR / ".fa_key"
USERS_FILE    = DATA_DIR / "users.json"
API_PORT      = 8788

# Real sound files (same names/location as the desktop 3.x build) — drop them
# beside this script and they're used automatically. Missing files just mean
# that cue falls back to the synthesized beep() in the browser.
SOUND_FILES = {
    "startup":      APP_DIR / "sound_startup.mp3",
    "camera_start": APP_DIR / "sound_camera_start.wav",
    "error":        APP_DIR / "sound_error.wav",
    "success":      APP_DIR / "sound_success.mp3",
    "negative":     APP_DIR / "sound_negative.mp3",
    "export":       APP_DIR / "sound_export.mp3",
}

COLUMNS = [
    "Sl.No", "Roll No", "UID", "Name", "Grade", "House",
    "Parent Phone", "Parent Email", "No. of Days Present",
    "XP", "Multiplier", "Streak",
]
HOUSES = ["Red", "Blue", "Green", "Orange"]

OLLAMA_URL   = os.getenv("FACEATTEND_OLLAMA_URL", "http://localhost:11434")
OLLAMA_MODEL = os.getenv("FACEATTEND_OLLAMA_MODEL", "qwen2.5:0.5b")


def ensure_dirs() -> None:
    for d in (DATA_DIR, FACES_DIR, UNKNOWNS_DIR, LOGS_DIR, BACKUPS_DIR,
              EXPORTS_DIR, BOARD_DIR, NOTES_DIR, CERTS_DIR, QR_DIR,
              SLIPS_DIR, EVIDENCE_DIR):
        d.mkdir(parents=True, exist_ok=True)


def safe_int(v, d=0) -> int:
    try:
        return int(float(v))
    except Exception:
        return d


def safe_float(v, d=0.0) -> float:
    try:
        return float(v)
    except Exception:
        return d


# ══════════════════════════════════════════════════════════════════════════
#  SETTINGS
# ══════════════════════════════════════════════════════════════════════════
DEFAULT_SETTINGS = {
    "late_time": "08:30",
    "school_name": "",
    "threshold": 0.45,
    # How many consecutive sightings must agree before someone is marked.
    "confirm_frames": 2,
    # How far ahead of the runner-up the best match must be, before a face
    # is accepted. Off by default: in testing it held back roughly one
    # sighting in five while preventing no wrong marks, because encodings
    # are far enough apart that the nearest match is almost always right.
    # Turn it up (0.05-0.08) only if two students are genuinely being
    # confused — the recognition test will tell you if they are.
    "match_margin": 0.0,
    "xp_ontime": 5,
    "xp_late": 2,
    "smtp_user": "",
    "smtp_pass": "",
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    # ── carried over from 3.2 ────────────────────────────────────────
    "liveness_check": True,
    "camera_index": 0,
    # Each entry is {"name": "Main gate", "source": "0" or an rtsp:// URL}.
    # An empty list means "just use camera_index", which is the old behaviour.
    "cameras": [],
    # ── school timetable ─────────────────────────────────────────────
    # enforce=False keeps the old permissive behaviour: any scan counts as
    # present. Turn it on and the windows below decide the status instead.
    "schedule": {
        "enforce": False,
        "days": {
            "mon": {"active": True, "start": "07:30", "late": "08:30",
                    "half": "11:00", "close": "15:30"},
            "tue": {"active": True, "start": "07:30", "late": "08:30",
                    "half": "11:00", "close": "15:30"},
            "wed": {"active": True, "start": "07:30", "late": "08:30",
                    "half": "11:00", "close": "15:30"},
            "thu": {"active": True, "start": "07:30", "late": "08:30",
                    "half": "11:00", "close": "15:30"},
            "fri": {"active": True, "start": "07:30", "late": "08:30",
                    "half": "11:00", "close": "15:30"},
            "sat": {"active": False, "start": "08:00", "late": "08:45",
                    "half": "10:30", "close": "12:30"},
            "sun": {"active": False, "start": "08:00", "late": "08:45",
                    "half": "10:30", "close": "12:30"},
        },
        "holidays": [],          # ISO dates, e.g. "2026-10-02"
    },
    "xp_half": 1,
    "auto_backup": True,
    "backup_keep": 14,
    "low_attendance_pct": 85.0,
    "keep_scan_photos": True,
    "scan_photo_days": 30,
    "high_contrast": False,
    "api_server": False,
    "voice_feedback": True,
    "encrypt_data": False,
    "cert_min_pct": 90.0,
    "auto_notify_absent": False,
    "absent_notify_time": "10:00",
    "whatsapp_enabled": False,
}


class Settings:
    def __init__(self):
        self.data = dict(DEFAULT_SETTINGS)
        self.load()

    def load(self):
        try:
            if SETTINGS_FILE.exists():
                self.data.update(json.loads(SETTINGS_FILE.read_text("utf-8")))
        except Exception:
            pass

    def save(self, patch: dict | None = None):
        if patch:
            for k, v in patch.items():
                if k not in DEFAULT_SETTINGS:
                    continue
                # The password box is deliberately left blank when the page
                # loads, so an empty value here means "unchanged", not
                # "clear it". Without this the app password was wiped every
                # time settings were saved, and had to be typed again daily.
                if k == "smtp_pass" and not str(v).strip():
                    continue
                self.data[k] = v
        try:
            SETTINGS_FILE.write_text(json.dumps(self.data, indent=2), "utf-8")
        except Exception:
            pass

    def __getitem__(self, k):
        return self.data.get(k, DEFAULT_SETTINGS.get(k))


# ══════════════════════════════════════════════════════════════════════════
#  STORE  — same Excel schema as FaceAttend 3.x
# ══════════════════════════════════════════════════════════════════════════



# ══════════════════════════════════════════════════════════════════════════
#  ATTENDANCE TIMING
#
#  One cutoff was never enough. A student walking in at 3pm was recorded as
#  "present" exactly like one who arrived at 8am, a scan on a Sunday created
#  a school day that dragged everyone's percentage down, and "absent" only
#  ever meant "has not scanned yet" — so at 8am the whole school read absent.
#
#  The day is now divided into windows:
#
#      before start   -> too_early   (register not open)
#      start .. late  -> present
#      late  .. half  -> late
#      half  .. close -> half_day
#      after close    -> absent      (register closed)
#      inactive day / holiday -> not a school day at all
#
#  This function is the single source of truth; the scanner, the manual
#  mark button and the reports all ask it rather than deciding for
#  themselves.
# ══════════════════════════════════════════════════════════════════════════
WEEKDAY_KEYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]

STATUS_LABEL = {
    "present":   "On time",
    "late":      "Late",
    "half_day":  "Half day",
    "absent":    "Absent",
    "too_early": "Before school",
    "holiday":   "Holiday",
    "closed":    "Not a school day",
}


def _hhmm(value: str, default: str) -> tuple[int, int]:
    try:
        hh, mm = str(value or default).split(":")
        return max(0, min(23, int(hh))), max(0, min(59, int(mm)))
    except Exception:
        hh, mm = default.split(":")
        return int(hh), int(mm)


def day_schedule(when: datetime | None = None) -> dict:
    """The timetable that applies on a given date."""
    when = when or datetime.now()
    sched = settings["schedule"] or {}
    days = sched.get("days") or {}
    key = WEEKDAY_KEYS[when.weekday()]
    cfg = dict(days.get(key) or {})
    cfg.setdefault("active", True)
    cfg.setdefault("start", "07:30")
    # Fall back to the old single setting so upgrading changes nothing.
    cfg.setdefault("late", settings["late_time"] or "08:30")
    cfg.setdefault("half", "11:00")
    cfg.setdefault("close", "15:30")
    cfg["key"] = key
    cfg["holiday"] = when.strftime("%Y-%m-%d") in (sched.get("holidays") or [])
    cfg["enforce"] = bool(sched.get("enforce"))
    return cfg


def is_school_day(when: datetime | None = None) -> bool:
    cfg = day_schedule(when)
    if not cfg["enforce"]:
        return True                       # permissive mode: every day counts
    return bool(cfg["active"]) and not cfg["holiday"]


def attendance_status(when: datetime | None = None) -> dict:
    """
    Work out what a scan right now should count as.

    Returns status, a human label, the XP to award, and whether marking
    is allowed at all.
    """
    when = when or datetime.now()
    cfg = day_schedule(when)

    xp_on = safe_int(settings["xp_ontime"], 5)
    xp_late = safe_int(settings["xp_late"], 2)
    xp_half = safe_int(settings["xp_half"], 1)

    if not cfg["enforce"]:
        # Old behaviour, kept as the default so nothing breaks mid-demo:
        # one cutoff, and any scan counts.
        h, m = _hhmm(cfg["late"], "08:30")
        late = when > when.replace(hour=h, minute=m, second=0, microsecond=0)
        return {"status": "late" if late else "present",
                "label": "Late" if late else "On time",
                "xp": xp_late if late else xp_on,
                "can_mark": True, "school_day": True, "enforced": False,
                "window": f"cutoff {cfg['late']}"}

    if cfg["holiday"]:
        return {"status": "holiday", "label": "Holiday", "xp": 0,
                "can_mark": False, "school_day": False, "enforced": True,
                "window": "holiday"}
    if not cfg["active"]:
        return {"status": "closed", "label": "Not a school day", "xp": 0,
                "can_mark": False, "school_day": False, "enforced": True,
                "window": cfg["key"]}

    at = lambda v, d: when.replace(hour=_hhmm(v, d)[0], minute=_hhmm(v, d)[1],
                                   second=0, microsecond=0)
    t_start, t_late = at(cfg["start"], "07:30"), at(cfg["late"], "08:30")
    t_half, t_close = at(cfg["half"], "11:00"), at(cfg["close"], "15:30")

    if when < t_start:
        return {"status": "too_early", "label": "Before school opens", "xp": 0,
                "can_mark": False, "school_day": True, "enforced": True,
                "window": f"opens {cfg['start']}"}
    if when <= t_late:
        return {"status": "present", "label": "On time", "xp": xp_on,
                "can_mark": True, "school_day": True, "enforced": True,
                "window": f"{cfg['start']}–{cfg['late']}"}
    if when <= t_half:
        return {"status": "late", "label": "Late", "xp": xp_late,
                "can_mark": True, "school_day": True, "enforced": True,
                "window": f"{cfg['late']}–{cfg['half']}"}
    if when <= t_close:
        return {"status": "half_day", "label": "Half day", "xp": xp_half,
                "can_mark": True, "school_day": True, "enforced": True,
                "window": f"{cfg['half']}–{cfg['close']}"}
    return {"status": "absent", "label": "Register closed", "xp": 0,
            "can_mark": False, "school_day": True, "enforced": True,
            "window": f"closed {cfg['close']}"}


def working_days(frm: str = "", to: str = "") -> list[str]:
    """
    School days in a range, honouring the timetable.

    Reports used to divide by "days that had any scan at all", so a single
    weekend test permanently lowered everyone's percentage.
    """
    sched = settings["schedule"] or {}
    if not sched.get("enforce"):
        return []                        # caller falls back to observed days
    days = sched.get("days") or {}
    holidays = set(sched.get("holidays") or [])
    try:
        start = datetime.strptime(frm, "%Y-%m-%d").date() if frm else None
        end = datetime.strptime(to, "%Y-%m-%d").date() if to else date.today()
    except Exception:
        return []
    if start is None:
        return []
    out, cur = [], start
    while cur <= end:
        key = WEEKDAY_KEYS[cur.weekday()]
        iso = cur.strftime("%Y-%m-%d")
        if (days.get(key) or {}).get("active", True) and iso not in holidays:
            out.append(iso)
        cur += timedelta(days=1)
    return out


def audit_rows() -> list[dict]:
    try:
        if AUDIT_LOG.exists():
            with open(AUDIT_LOG, newline="", encoding="utf-8") as f:
                return list(csv.DictReader(f))
    except Exception:
        pass
    return []


def audit_last_hash() -> str:
    rows = audit_rows()
    return rows[-1].get("hash", "") if rows else "genesis"


def audit_verify() -> dict:
    """Recompute the chain and report the first row that does not match."""
    rows = audit_rows()
    prev = "genesis"
    for i, r in enumerate(rows):
        if "hash" not in r or r.get("hash") is None:
            return {"ok": False, "checked": i,
                    "problem": f"Row {i + 1} predates hashing.", "legacy": True}
        body = [r.get(k, "") or "" for k in
                ("date", "time", "uid", "name", "status", "xp", "reason", "actor")]
        want = hashlib.sha256((prev + "|" + "|".join(body)).encode()).hexdigest()[:16]
        if want != r.get("hash"):
            return {"ok": False, "checked": i + 1,
                    "problem": f"Row {i + 1} ({r.get('name')}, {r.get('date')}) "
                               "does not match the chain — it may have been edited."}
        prev = r["hash"]
    return {"ok": True, "checked": len(rows), "problem": ""}


def validate_student(data: dict, students: list[dict], editing_uid: str = "") -> list[str]:
    """
    Check a student record before it is saved. Returns a list of problems;
    an empty list means the record is fine.

    Without this, duplicate roll numbers and malformed emails quietly pile up
    and only surface months later when a report looks wrong.
    """
    problems = []
    name = str(data.get("Name", "")).strip()
    roll = str(data.get("Roll No", "")).strip()
    email = str(data.get("Parent Email", "")).strip()
    phone = re.sub(r"[^\d]", "", str(data.get("Parent Phone", "")))
    house = str(data.get("House", "")).strip()

    if len(name) < 2:
        problems.append("Name must be at least 2 characters.")
    if len(name) > 60:
        problems.append("Name is unusually long — check for a typo.")
    if email and not re.match(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$", email):
        problems.append(f"'{email}' does not look like a valid email address.")
    if phone and not (7 <= len(phone) <= 15):
        problems.append("Phone number should be 7 to 15 digits.")
    if house and house not in HOUSES:
        problems.append(f"House must be one of: {', '.join(HOUSES)}.")

    for other in students:
        if str(other.get("UID")) == str(editing_uid):
            continue          # don't compare a record against itself
        o_roll = str(other.get("Roll No", "")).strip()
        if roll and o_roll and roll.lower() == o_roll.lower():
            problems.append(f"Roll number '{roll}' already belongs to "
                            f"{other.get('Name')}.")
            break

    same_name = [o for o in students
                 if str(o.get("Name", "")).strip().lower() == name.lower()
                 and str(o.get("UID")) != str(editing_uid)]
    if same_name and not roll:
        problems.append(f"A student called '{name}' already exists — "
                        "add a roll number to tell them apart.")
    return problems


class Store:
    def __init__(self):
        self.lock = threading.Lock()
        self.ensure_workbook()

    def ensure_workbook(self):
        if load_workbook is None:
            return
        if EXCEL_FILE.exists():
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Students"
        ws.append(COLUMNS)
        for i, name in enumerate(COLUMNS, 1):
            c = ws.cell(1, i)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", start_color="2F6BFF")
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(i)].width = max(11, len(name) + 4)
        wb.save(EXCEL_FILE)

    def load_students(self) -> list[dict]:
        if load_workbook is None or not EXCEL_FILE.exists():
            return []
        try:
            with self.lock:
                wb = load_workbook(EXCEL_FILE, data_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
        except Exception:
            return []
        if not rows:
            return []
        header = [str(h) if h is not None else "" for h in rows[0]]
        out = []
        for r in rows[1:]:
            if r is None or all(v is None for v in r):
                continue
            rec = dict(zip(header, r))
            s = {c: rec.get(c) for c in COLUMNS}
            if not s.get("UID") and not s.get("Name"):
                continue
            s["UID"] = str(s.get("UID") or "").strip()
            s["Name"] = str(s.get("Name") or "").strip()
            s["House"] = str(s.get("House") or "").strip() or "Blue"
            for k in ("No. of Days Present", "XP", "Streak"):
                s[k] = safe_int(s.get(k), 0)
            s["Multiplier"] = safe_float(s.get("Multiplier"), 1.0)
            out.append(s)
        return out

    def write_all(self, students: list[dict]):
        if Workbook is None:
            return
        with self.lock:
            wb = Workbook()
            ws = wb.active
            ws.title = "Students"
            ws.append(COLUMNS)
            for i, name in enumerate(COLUMNS, 1):
                c = ws.cell(1, i)
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", start_color="2F6BFF")
                c.alignment = Alignment(horizontal="center", vertical="center")
                ws.column_dimensions[get_column_letter(i)].width = max(11, len(name) + 4)
            thin = Side(style="thin", color="D8DEE9")
            for idx, s in enumerate(students, 1):
                s["Sl.No"] = idx
                ws.append([s.get(c, "") for c in COLUMNS])
                for col in range(1, len(COLUMNS) + 1):
                    ws.cell(idx + 1, col).border = Border(thin, thin, thin, thin)
            ws.freeze_panes = "A2"
            # Write to a temp file in the same folder, then rename. os.replace
            # is atomic, so killing the app mid-save can no longer leave a
            # half-written roster — the old file survives intact until the new
            # one is complete.
            tmp = EXCEL_FILE.with_suffix(".xlsx.tmp")
            try:
                wb.save(tmp)
                os.replace(tmp, EXCEL_FILE)
            except Exception:
                try:
                    tmp.unlink(missing_ok=True)
                except Exception:
                    pass
                raise

    def next_uid(self, students) -> str:
        nums = []
        for s in students:
            m = re.match(r"F(\d+)", str(s.get("UID", "")))
            if m:
                nums.append(int(m.group(1)))
        return f"F{(max(nums) + 1) if nums else 1:03d}"

    def add(self, data: dict) -> str:
        students = self.load_students()
        uid = self.next_uid(students)
        rec = {c: "" for c in COLUMNS}
        rec.update({
            "UID": uid,
            "Name": data.get("Name", "").strip(),
            "Roll No": data.get("Roll No", ""),
            "Grade": data.get("Grade", ""),
            "House": data.get("House") or "Blue",
            "Parent Phone": data.get("Parent Phone", ""),
            "Parent Email": data.get("Parent Email", ""),
            "No. of Days Present": 0, "XP": 0, "Multiplier": 1.0, "Streak": 0,
        })
        students.append(rec)
        self.write_all(students)
        return uid

    def update(self, uid: str, data: dict) -> bool:
        students = self.load_students()
        hit = False
        for s in students:
            if str(s.get("UID")) == str(uid):
                for k in ("Name", "Roll No", "Grade", "House", "Parent Phone", "Parent Email"):
                    if k in data:
                        s[k] = data[k]
                hit = True
        if hit:
            self.write_all(students)
        return hit

    def remove(self, uid: str) -> bool:
        students = self.load_students()
        keep = [s for s in students if str(s.get("UID")) != str(uid)]
        if len(keep) == len(students):
            return False
        self.write_all(keep)
        for p in FACES_DIR.glob(f"{uid}.*"):
            try:
                p.unlink()
            except Exception:
                pass
        return True

    def award(self, uid: str, xp: int) -> dict | None:
        students = self.load_students()
        out = None
        for s in students:
            if str(s.get("UID")) == str(uid):
                s["No. of Days Present"] = safe_int(s.get("No. of Days Present")) + 1
                s["XP"] = safe_int(s.get("XP")) + xp
                s["Streak"] = safe_int(s.get("Streak")) + 1
                out = dict(s)
        if out:
            self.write_all(students)
        return out

    def house_stats(self, present_uids: set, students: list | None = None) -> dict:
        # Callers on a hot path can pass the roster they already have,
        # avoiding another full Excel read.
        if students is None:
            students = self.load_students()
        out = {}
        for h in HOUSES:
            members = [s for s in students if (s.get("House") or "Blue") == h]
            pres = [s for s in members if str(s.get("UID")) in present_uids]
            out[h] = {
                "students": len(members),
                "present": len(pres),
                "pct": (len(pres) / len(members) * 100) if members else 0.0,
            }
        return out


# ══════════════════════════════════════════════════════════════════════════
#  BADGES                                              (ported from 3.2)
# ══════════════════════════════════════════════════════════════════════════
class Badges:
    """Achievement tracking. Same six badges and thresholds as 3.2."""

    DEFS = {
        "first_steps":  {"name": "First Steps",  "icon": "\U0001F463",
                         "desc": "Attend for the first time",     "req_days": 1},
        "week_warrior": {"name": "Week Warrior", "icon": "\u2694",
                         "desc": "Attend 5 days",                 "req_days": 5},
        "month_master": {"name": "Month Master", "icon": "\U0001F451",
                         "desc": "Attend 20 days",                "req_days": 20},
        "early_bird":   {"name": "Early Bird",   "icon": "\U0001F426",
                         "desc": "Arrive early 10 times",         "req_early": 10},
        "centurion":    {"name": "Centurion",    "icon": "\U0001F3DB",
                         "desc": "Attend 100 days",               "req_days": 100},
        "house_hero":   {"name": "House Hero",   "icon": "\U0001F3C6",
                         "desc": "Your house reaches 90%",        "req_house_pct": 90.0},
    }

    def __init__(self, store):
        self.store = store
        self._awards: dict[str, list[str]] = {}
        self.load()

    def load(self):
        try:
            if BADGE_FILE.exists():
                self._awards = json.loads(BADGE_FILE.read_text("utf-8"))
        except Exception:
            self._awards = {}

    def save(self):
        try:
            BADGE_FILE.parent.mkdir(parents=True, exist_ok=True)
            BADGE_FILE.write_text(json.dumps(self._awards, indent=2), "utf-8")
        except Exception:
            pass

    def for_student(self, uid: str) -> list[dict]:
        earned = self._awards.get(str(uid), [])
        return [dict(self.DEFS[k], key=k) for k in self.DEFS if k in earned]

    def all_with_progress(self, student: dict) -> list[dict]:
        """Every badge, marked earned or not, with progress toward it."""
        uid = str(student.get("UID", ""))
        earned = self._awards.get(uid, [])
        days = safe_int(student.get("No. of Days Present"))
        out = []
        for key, b in self.DEFS.items():
            item = dict(b, key=key, earned=key in earned)
            if "req_days" in b:
                item["progress"] = min(100, days / b["req_days"] * 100)
                item["detail"] = f"{days}/{b['req_days']} days"
            elif "req_early" in b:
                n = self.count_early(uid)
                item["progress"] = min(100, n / b["req_early"] * 100)
                item["detail"] = f"{n}/{b['req_early']} early"
            else:
                item["progress"] = 100 if key in earned else 0
                item["detail"] = "House at 90%"
            out.append(item)
        return out

    def count_early(self, uid: str) -> int:
        n = 0
        try:
            if AUDIT_LOG.exists():
                with open(AUDIT_LOG, newline="", encoding="utf-8") as f:
                    for row in csv.DictReader(f):
                        if row.get("uid") == str(uid) and row.get("status") == "ontime":
                            n += 1
        except Exception:
            pass
        return n

    def check_and_award(self, uid: str, student: dict,
                        present_uids: set | None = None) -> list[dict]:
        uid = str(uid)
        earned = self._awards.setdefault(uid, [])
        days = safe_int(student.get("No. of Days Present"))
        new = []
        for key, b in self.DEFS.items():
            if key in earned:
                continue
            hit = False
            if "req_days" in b:
                hit = days >= b["req_days"]
            elif "req_early" in b:
                hit = self.count_early(uid) >= b["req_early"]
            elif "req_house_pct" in b:
                house = str(student.get("House") or "").strip()
                if house:
                    stats = self.store.house_stats(present_uids or set())
                    hit = stats.get(house, {}).get("pct", 0) >= b["req_house_pct"]
            if hit:
                earned.append(key)
                new.append(dict(b, key=key))
        if new:
            self.save()
        return new


# ══════════════════════════════════════════════════════════════════════════
#  CERTIFICATES                                        (ported from 3.2)
# ══════════════════════════════════════════════════════════════════════════
def generate_certificate(path, name, roll, grade, pct, days, school) -> bool:
    """Landscape A4 attendance certificate. Same layout as 3.2."""
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as rl_canvas
    except Exception:
        return False

    W, H = landscape(A4)
    navy = colors.HexColor("#1e3a5f")
    gold = colors.HexColor("#c9a227")
    cream = colors.HexColor("#faf6ee")
    ink = colors.HexColor("#243b53")

    try:
        c = rl_canvas.Canvas(str(path), pagesize=landscape(A4))
        c.setFillColor(cream)
        c.rect(0, 0, W, H, fill=1, stroke=0)

        c.setStrokeColor(navy)
        c.setLineWidth(6)
        c.rect(22, 22, W - 44, H - 44, fill=0, stroke=1)
        c.setStrokeColor(gold)
        c.setLineWidth(1.6)
        c.rect(34, 34, W - 68, H - 68, fill=0, stroke=1)

        for cx, cy in ((52, 52), (W - 52, 52), (52, H - 52), (W - 52, H - 52)):
            c.setFillColor(gold)
            c.circle(cx, cy, 7, fill=1, stroke=0)

        c.setFillColor(navy)
        c.setFont("Helvetica-Bold", 34)
        c.drawCentredString(W / 2, H - 118, "CERTIFICATE OF ATTENDANCE")
        c.setStrokeColor(gold)
        c.setLineWidth(2)
        c.line(W / 2 - 168, H - 132, W / 2 + 168, H - 132)

        c.setFillColor(ink)
        c.setFont("Helvetica", 14)
        c.drawCentredString(W / 2, H - 172, "This is to certify that")

        c.setFillColor(navy)
        c.setFont("Helvetica-Bold", 30)
        c.drawCentredString(W / 2, H - 216, str(name))
        c.setStrokeColor(gold)
        c.setLineWidth(1)
        c.line(W / 2 - 190, H - 228, W / 2 + 190, H - 228)

        c.setFillColor(ink)
        c.setFont("Helvetica", 13)
        sub = f"Roll No. {roll}" + (f"   \u00b7   Grade {grade}" if grade else "")
        c.drawCentredString(W / 2, H - 252, sub)

        c.setFont("Helvetica", 14)
        c.drawCentredString(W / 2, H - 296,
                            "has achieved an outstanding attendance record of")
        c.setFillColor(gold)
        c.setFont("Helvetica-Bold", 46)
        c.drawCentredString(W / 2, H - 352, f"{pct:.1f}%")
        c.setFillColor(ink)
        c.setFont("Helvetica", 13)
        c.drawCentredString(W / 2, H - 378, f"across {days} school days")

        c.setFont("Helvetica-Oblique", 11)
        c.drawString(78, 96, datetime.now().strftime("%d %B %Y"))
        c.setStrokeColor(ink)
        c.setLineWidth(0.8)
        c.line(78, 112, 250, 112)
        c.setFont("Helvetica", 9)
        c.drawString(78, 82, "Date")

        c.line(W - 250, 112, W - 78, 112)
        c.setFont("Helvetica-Oblique", 11)
        c.drawRightString(W - 78, 96, str(school or "School"))
        c.setFont("Helvetica", 9)
        c.drawRightString(W - 78, 82, "Authorised signature")

        c.showPage()
        c.save()
        return True
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════════════
#  QR CODES                                            (ported from 3.2)
# ══════════════════════════════════════════════════════════════════════════
def generate_student_qr(uid: str, name: str, roll: str):
    """PNG QR encoding the student's identity. Returns a path or None."""
    try:
        import qrcode
    except Exception:
        return None
    try:
        QR_DIR.mkdir(parents=True, exist_ok=True)
        payload = json.dumps({"uid": uid, "name": name, "roll": roll,
                              "app": "FaceAttend"})
        img = qrcode.make(payload)
        out = QR_DIR / f"{uid}.png"
        img.save(out)
        return out
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════════════════
#  ENCRYPTION                                          (ported from 3.2)
# ══════════════════════════════════════════════════════════════════════════
class Vault:
    """Optional at-rest encryption for the roster, using Fernet."""

    def __init__(self):
        self._key = None

    def key(self):
        if self._key is not None:
            return self._key
        try:
            from cryptography.fernet import Fernet
        except Exception:
            return None
        try:
            if KEY_FILE.exists():
                self._key = KEY_FILE.read_bytes()
            else:
                self._key = Fernet.generate_key()
                KEY_FILE.parent.mkdir(parents=True, exist_ok=True)
                KEY_FILE.write_bytes(self._key)
                try:
                    os.chmod(KEY_FILE, 0o600)
                except Exception:
                    pass
        except Exception:
            return None
        return self._key

    def available(self) -> bool:
        try:
            import cryptography  # noqa: F401
            return True
        except Exception:
            return False

    def encrypt_file(self, path: Path) -> bool:
        k = self.key()
        if k is None or not path.exists():
            return False
        try:
            from cryptography.fernet import Fernet
            data = path.read_bytes()
            if data.startswith(b"gAAAAA"):
                return True                       # already encrypted
            path.with_suffix(path.suffix + ".enc").write_bytes(Fernet(k).encrypt(data))
            return True
        except Exception:
            return False

    def decrypt_file(self, path: Path) -> bool:
        k = self.key()
        if k is None or not path.exists():
            return False
        try:
            from cryptography.fernet import Fernet
            out = Fernet(k).decrypt(path.read_bytes())
            path.with_suffix("").write_bytes(out)
            return True
        except Exception:
            return False


# ══════════════════════════════════════════════════════════════════════════
#  CALENDAR HEATMAP                                    (ported from 3.2)
# ══════════════════════════════════════════════════════════════════════════
def calendar_data(year: int, month: int, total_students: int) -> dict:
    """Per-day attendance counts for a month, read from the audit log."""
    import calendar as _cal
    seen: dict[str, set] = {}
    try:
        if AUDIT_LOG.exists():
            with open(AUDIT_LOG, newline="", encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    d = row.get("date")
                    if d:
                        seen.setdefault(d, set()).add(row.get("uid"))
    except Exception:
        pass
    days_in = _cal.monthrange(year, month)[1]
    first_wd = _cal.monthrange(year, month)[0]      # 0 = Monday
    days = []
    for d in range(1, days_in + 1):
        key = f"{year:04d}-{month:02d}-{d:02d}"
        n = len(seen.get(key, ()))
        days.append({
            "day": d,
            "count": n,
            "pct": (n / total_students * 100) if total_students else 0.0,
            "weekend": _cal.weekday(year, month, d) >= 5,
        })
    return {"year": year, "month": month, "first_weekday": first_wd,
            "month_name": _cal.month_name[month], "days": days}


# ══════════════════════════════════════════════════════════════════════════
#  LIVENESS                                            (ported from 3.2)
# ══════════════════════════════════════════════════════════════════════════
def liveness_ok(frame, box) -> bool:
    """
    Cheap anti-spoof: a photo held to the camera is flatter and less
    textured than a real face. Uses Laplacian variance on the face crop.
    Not bulletproof, but it stops the obvious phone-screen attack.
    """
    if cv2 is None or np is None:
        return True
    try:
        top, right, bottom, left = box
        crop = frame[max(0, top):bottom, max(0, left):right]
        if crop.size == 0:
            return True
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        sharpness = cv2.Laplacian(gray, cv2.CV_64F).var()
        colour_spread = float(np.std(crop.reshape(-1, 3), axis=0).mean())
        return sharpness > 22.0 and colour_spread > 8.0
    except Exception:
        return True


# ══════════════════════════════════════════════════════════════════════════
#  SCAN STATE  — camera + recognition thread
# ══════════════════════════════════════════════════════════════════════════
class Scanner:
    def __init__(self, store: Store, settings: Settings):
        self.store = store
        self.settings = settings
        self.cap = None
        self.running = False
        self.thread = None
        self.frame = None
        self.frame_lock = threading.Lock()
        self.last = {"stamp": None, "name": None, "uid": None, "xp": 0, "late": False, "time": ""}
        self.scanned: set[str] = set()
        self.unknown_count = 0
        self.liveness_rejects = 0
        self.ambiguous = 0
        self.last_badges: list[str] = []
        # Writes are drained by a background worker so the camera loop never
        # blocks on disk I/O.
        self.boxes: list = []          # last known face boxes, drawn every frame
        self.active_source = ""
        self.last_status = ""
        self._latest = None            # newest raw frame, for recognition
        self._last_unknown = 0.0
        import queue as _q
        self._persist: "_q.Queue" = _q.Queue()
        self._writer = threading.Thread(target=self._persist_loop, daemon=True)
        self._writer.start()
        self.encodings: dict[str, object] = {}
        self.load_scans()
        self.load_encodings()

    # ── persistence ───────────────────────────────────────────────
    def load_scans(self):
        try:
            if SCANS_FILE.exists():
                blob = pickle.loads(SCANS_FILE.read_bytes())
                if blob.get("date") == str(date.today()):
                    self.scanned = set(blob.get("uids", []))
        except Exception:
            pass

    def save_scans(self):
        try:
            SCANS_FILE.write_bytes(pickle.dumps(
                {"date": str(date.today()), "uids": list(self.scanned)}))
        except Exception:
            pass

    def load_encodings(self):
        """
        Load face encodings, converting older layouts on the way in.

        FaceAttend 3.x stored two parallel lists:
            {"uids": [...], "encodings": [...]}
        The web build uses {uid: [encoding, ...]}. Reading a 3.x file
        without converting produced a dict whose "student IDs" were the
        literal strings "uids" and "encodings", so nothing ever matched
        and attendance was never marked automatically.
        """
        self.encodings = {}
        try:
            if not ENCODINGS_FILE.exists():
                return
            raw = pickle.loads(ENCODINGS_FILE.read_bytes())
        except Exception:
            return

        if not isinstance(raw, dict):
            return

        # ── legacy 3.x parallel-list format ──────────────────────────
        if "uids" in raw and "encodings" in raw:
            uids = list(raw.get("uids") or [])
            vecs = list(raw.get("encodings") or [])
            merged: dict = {}
            for uid, vec in zip(uids, vecs):
                merged.setdefault(str(uid), []).append(vec)
            self.encodings = merged
            print(f"  migrated {len(uids)} face encoding(s) from the 3.x format "
                  f"({len(merged)} student(s))")
            self.save_encodings()          # rewrite in the current layout
            return

        # ── current format; normalise single encodings into lists ────
        for uid, val in raw.items():
            self.encodings[str(uid)] = val if isinstance(val, list) else [val]

    def save_encodings(self):
        try:
            ENCODINGS_FILE.write_bytes(pickle.dumps(self.encodings))
        except Exception:
            pass

    def learn(self, uid: str, image_path: Path, append: bool = False) -> bool:
        """
        Compute and store a face encoding for one student.

        With append=True the new encoding is added alongside any existing
        ones rather than replacing them. Several photos from different
        angles and lighting conditions raise recognition accuracy a great
        deal — a single front-on shot fails as soon as someone turns their
        head or the light changes.
        """
        if face_recognition is None:
            return False
        try:
            img = face_recognition.load_image_file(str(image_path))
            encs = face_recognition.face_encodings(img)
            if not encs:
                return False
            existing = self.encodings.get(uid)
            if append and existing is not None:
                current = existing if isinstance(existing, list) else [existing]
                current.append(encs[0])
                self.encodings[uid] = current[-6:]      # keep the 6 most recent
            else:
                self.encodings[uid] = [encs[0]]
            self.save_encodings()
            return True
        except Exception:
            return False

    # ── camera ────────────────────────────────────────────────────
    @staticmethod
    def _backends():
        """
        Capture backends to try, best first.

        Windows really wants DirectShow or MSMF — the default backend there
        can take 5+ seconds to open a webcam, or fail outright. On Linux the
        default (V4L2) is correct.
        """
        if cv2 is None:
            return []
        if IS_WINDOWS:
            return [getattr(cv2, "CAP_DSHOW", 700),
                    getattr(cv2, "CAP_MSMF", 1400), 0]
        if IS_MAC:
            return [getattr(cv2, "CAP_AVFOUNDATION", 1200), 0]
        return [0]

    def _flat_encodings(self):
        """
        Expand the per-student encoding lists into two parallel arrays.

        A student may now have several encodings; each becomes its own entry
        pointing back at the same UID, so a match against any one of them
        identifies the student.
        """
        uids, vecs = [], []
        for uid, enc in self.encodings.items():
            for e in (enc if isinstance(enc, list) else [enc]):
                uids.append(uid)
                vecs.append(e)
        return uids, vecs

    def _sources(self) -> list[dict]:
        """
        Where to look for video.

        A plain number is a local USB camera; anything else is treated as a
        stream URL, which is what lets FaceAttend read IP cameras at a gate
        instead of only a webcam plugged into this machine.
        """
        cams = self.settings["cameras"] or []
        if cams:
            return [{"name": c.get("name") or f"Camera {i+1}",
                     "source": str(c.get("source", "")).strip()}
                    for i, c in enumerate(cams) if str(c.get("source", "")).strip()]
        return [{"name": "Default camera",
                 "source": str(safe_int(self.settings["camera_index"], 0))}]

    def open_stream(self, source: str):
        """Open one source, USB index or network URL, and return the capture."""
        if cv2 is None:
            return None
        if source.isdigit():
            for backend in self._backends():
                try:
                    cap = (cv2.VideoCapture(int(source), backend) if backend
                           else cv2.VideoCapture(int(source)))
                except Exception:
                    continue
                if cap.isOpened():
                    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
                    return cap
                cap.release()
            return None
        try:
            cap = cv2.VideoCapture(source)
            # A small buffer keeps the feed close to live; the default queues
            # several seconds of frames, so a student walks off before their
            # face is processed.
            try:
                cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
            except Exception:
                pass
            if cap.isOpened():
                return cap
            cap.release()
        except Exception:
            pass
        return None

    def open_camera(self) -> bool:
        if cv2 is None:
            return False
        for src in self._sources():
            cap = self.open_stream(src["source"])
            if cap is not None:
                self.cap = cap
                self.active_source = src["name"]
                return True
        for backend in self._backends():
            for idx in (0, 1, 2):
                try:
                    cap = cv2.VideoCapture(idx, backend) if backend else cv2.VideoCapture(idx)
                except Exception:
                    continue
                if cap.isOpened():
                    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
                    self.cap = cap
                    return True
                cap.release()
        return False

    def start(self) -> tuple[bool, str]:
        if self.running:
            return True, ""
        if cv2 is None:
            return False, "opencv-python is not installed"
        if not self.open_camera():
            return False, ("No camera found. Check it is plugged in and that no other "
                           "app (Zoom, Teams, Camera) is using it.")
        self.running = True
        self.boxes = []
        self.thread = threading.Thread(target=self._loop, daemon=True)
        self.thread.start()
        self.recog_thread = threading.Thread(target=self._recognise_loop, daemon=True)
        self.recog_thread.start()
        return True, ""

    def stop(self):
        self.running = False
        self.boxes = []
        time.sleep(0.25)
        if self.cap:
            try:
                self.cap.release()
            except Exception:
                pass
            self.cap = None

    def is_late(self) -> bool:
        """Kept for compatibility; the engine decides the real status now."""
        return attendance_status()["status"] in ("late", "half_day")

    # ── capture and recognition run on SEPARATE threads ──────────────
    #
    # They used to share one loop, so every face encoding (200-800 ms on
    # CPU) stalled the video. Now the capture thread only reads frames and
    # draws the most recent boxes, staying smooth at all times, while the
    # recognition thread works on a snapshot at its own pace.

    def _loop(self):
        """Capture thread: read, annotate with cached boxes, publish."""
        misses = 0
        while self.running and self.cap:
            ok, frame = self.cap.read()
            if not ok:
                # Network cameras drop out; a USB one rarely does. Rather than
                # dying silently, try to reopen after a run of failed reads.
                misses += 1
                if misses > 60:
                    misses = 0
                    try:
                        self.cap.release()
                    except Exception:
                        pass
                    self.cap = None
                    if self.open_camera():
                        continue
                    self.running = False
                    break
                time.sleep(0.05)
                continue
            misses = 0

            with self.frame_lock:
                self._latest = frame            # hand-off for recognition

            display = frame.copy()
            for box, label, colour in list(self.boxes):
                t, r, b, l = box
                try:
                    cv2.rectangle(display, (l, t), (r, b), colour, 3)
                    # Label sits above the box, as in the desktop build.
                    (tw, th), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX,
                                                  0.62, 2)
                    ly = max(th + 8, t)
                    cv2.rectangle(display, (l, ly - th - 10),
                                  (l + tw + 14, ly), colour, cv2.FILLED)
                    cv2.putText(display, label, (l + 7, ly - 6),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.62,
                                (10, 12, 20), 2, cv2.LINE_AA)
                except Exception:
                    pass

            with self.frame_lock:
                self.frame = display
            time.sleep(0.02)

    def _recognise_loop(self):
        """Recognition thread: slow work, never blocks the video."""
        students = {str(s["UID"]): s for s in self.store.load_students()}
        known_uids, known_vecs = self._flat_encodings()
        cooldown: dict[str, float] = {}
        streak: dict[str, int] = {}
        refreshed = time.time()

        while self.running:
            if face_recognition is None or not known_vecs:
                time.sleep(0.4)
                continue

            with self.frame_lock:
                frame = None if self._latest is None else self._latest.copy()
            if frame is None:
                time.sleep(0.08)
                continue

            try:
                # 0.25 keeps enough detail for a gate-distance face while
                # cutting the encoding cost roughly sixteen-fold.
                small = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
                rgb = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
                locs = face_recognition.face_locations(rgb)      # HOG, CPU-friendly
                encs = face_recognition.face_encodings(rgb, locs) if locs else []
            except Exception:
                time.sleep(0.15)
                continue

            thr = safe_float(self.settings["threshold"], 0.45)
            boxes = []

            for (top, right, bottom, left), enc in zip(locs, encs):
                try:
                    dists = face_recognition.face_distance(known_vecs, enc)
                    best = int(np.argmin(dists)) if len(dists) else -1
                    matched = best >= 0 and dists[best] <= thr
                    t, r, b, l = top * 4, right * 4, bottom * 4, left * 4

                    # ── margin test ──────────────────────────────────
                    # Being under the threshold is not enough. The best
                    # match must also be clearly ahead of the closest
                    # *different* student, otherwise an ambiguous face gets
                    # attributed to whoever it happens to sit nearest.
                    need_margin = safe_float(self.settings["match_margin"], 0.0)
                    if matched and need_margin > 0 and len(dists) > 1:
                        best_uid = known_uids[best]
                        others = [d for d, u_ in zip(dists, known_uids)
                                  if u_ != best_uid]
                        if others:
                            margin = float(min(others)) - float(dists[best])
                            if margin < need_margin:
                                self.ambiguous += 1
                                boxes.append(((t, r, b, l),
                                              "Not sure \u2014 too similar",
                                              (61, 176, 255)))
                                continue

                    if matched and self.settings["liveness_check"]:
                        if not liveness_ok(frame, (t, r, b, l)):
                            boxes.append(((t, r, b, l), "Liveness failed",
                                          (60, 90, 255)))
                            self.liveness_rejects += 1
                            continue

                    if matched:
                        uid = known_uids[best]
                        st = students.get(uid, {})
                        name = str(st.get("Name", uid))
                        now_t = time.time()

                        # ── confirmation ─────────────────────────────
                        # A single frame is not proof. Requiring the same
                        # person on consecutive passes removes the one-off
                        # mis-reads that a noisy frame produces.
                        need_seen = max(1, safe_int(self.settings["confirm_frames"], 2))
                        seen = streak.get(uid, 0) + 1
                        streak = {uid: seen}          # any other face resets it
                        if seen < need_seen and uid not in self.scanned:
                            boxes.append(((t, r, b, l),
                                          f"{name} \u2014 confirming {seen}/{need_seen}",
                                          (255, 190, 61)))
                            continue

                        if uid not in self.scanned and now_t - cooldown.get(uid, 0) > 6:
                            cooldown[uid] = now_t
                            info = attendance_status()
                            if not info["can_mark"]:
                                # Outside the register window: show why on the
                                # box rather than silently marking them present.
                                boxes.append(((t, r, b, l),
                                              f"{name} — {info['label']}",
                                              (61, 176, 255)))
                                continue
                            late = info["status"] in ("late", "half_day")
                            xp = info["xp"]
                            self.last_status = info["status"]
                            self.scanned.add(uid)
                            if st:
                                st["No. of Days Present"] = safe_int(
                                    st.get("No. of Days Present")) + 1
                                st["XP"] = safe_int(st.get("XP")) + xp
                                st["Streak"] = safe_int(st.get("Streak")) + 1
                            self.last = {
                                "stamp": time.time(), "name": name, "uid": uid,
                                "xp": xp, "late": late, "badges": [],
                                "status": info["status"], "label": info["label"],
                                "time": datetime.now().strftime("%H:%M:%S"),
                            }
                            # Written to the queue first, so a crash before
                            # the roster is saved does not lose the scan.
                            queue.add({"uid": uid, "xp": xp, "name": name,
                                       "status": info["status"],
                                       "date": str(date.today()),
                                       "at": datetime.now().isoformat(timespec="seconds")})
                            if self.settings["keep_scan_photos"]:
                                self._save_evidence(uid, frame, (t, r, b, l))
                            # Disk work happens on the writer thread.
                            self._persist.put((uid, xp, name, late))
                        # Show the match distance the way the 3.2 desktop
                        # build did — it makes it obvious at a glance how
                        # confident the match is.
                        score = float(dists[best])
                        tick_mark = " \u2713" if uid in self.scanned else ""
                        label = f"{name} {score:.2f}{tick_mark}"
                        boxes.append(((t, r, b, l), label, (168, 230, 46)))
                    else:
                        boxes.append(((t, r, b, l), "Unknown", (115, 93, 255)))
                        now_t = time.time()
                        if now_t - getattr(self, "_last_unknown", 0) > 8:
                            self._last_unknown = now_t
                            self.unknown_count += 1
                            try:
                                cv2.imwrite(str(UNKNOWNS_DIR /
                                    f"unknown_{datetime.now():%Y%m%d_%H%M%S}.jpg"),
                                    frame)
                            except Exception:
                                pass
                except Exception:
                    continue

            self.boxes = boxes

            # Pick up newly enrolled students every 20 seconds.
            if time.time() - refreshed > 20:
                refreshed = time.time()
                try:
                    students = {str(s["UID"]): s for s in self.store.load_students()}
                    self.load_encodings()
                    known_uids, known_vecs = self._flat_encodings()
                except Exception:
                    pass

            time.sleep(0.05)

    def _persist_loop(self):
        """Drain recognition events: Excel write, badges, audit row."""
        while True:
            try:
                uid, xp, name, late = self._persist.get()
            except Exception:
                return
            try:
                self.save_scans()
                rec = self.store.award(uid, xp)
                if rec is not None and badges is not None:
                    won = badges.check_and_award(uid, rec, self.scanned)
                    if won:
                        # Surface new badges on the next poll from the UI.
                        cur = self.last
                        if cur.get("uid") == uid:
                            cur["badges"] = [b["name"] for b in won]
                            cur["stamp"] = time.time()
                self._audit(uid, name, late, xp,
                            reason=getattr(self, "last_status", "") or "")
                queue.done({"uid": uid, "at": None})   # matched on uid
            except Exception:
                pass
            finally:
                try:
                    self._persist.task_done()
                except Exception:
                    pass


    def _save_evidence(self, uid, frame, box):
        """
        Keep the frame that produced a recognition.

        Settles any argument about whether a mark was genuine, and pairs with
        the hash-chained audit log. Pruned on a schedule so a term of scans
        does not fill the disk.
        """
        if cv2 is None:
            return
        try:
            EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
            t, r, b, l = box
            pad = 40
            h, w = frame.shape[:2]
            crop = frame[max(0, t - pad):min(h, b + pad),
                         max(0, l - pad):min(w, r + pad)]
            if crop.size == 0:
                crop = frame
            out = EVIDENCE_DIR / f"{uid}_{datetime.now():%Y%m%d_%H%M%S}.jpg"
            cv2.imwrite(str(out), crop, [int(cv2.IMWRITE_JPEG_QUALITY), 70])
        except Exception:
            pass

    @staticmethod
    def prune_evidence(days: int = 30):
        cutoff = time.time() - max(1, days) * 86400
        try:
            for f in EVIDENCE_DIR.glob("*.jpg"):
                if f.stat().st_mtime < cutoff:
                    f.unlink(missing_ok=True)
        except Exception:
            pass

    def _audit(self, uid, name, late, xp, action="mark", reason="", actor=""):
        """
        Append a row to the audit log, hash-chained to the row before it.

        Each row stores a hash of (previous hash + this row's contents). If
        anyone edits or deletes a past row the chain breaks and /api/audit
        reports it. For an attendance record that is the whole point — the
        log is only worth something if tampering is detectable.
        """
        try:
            new_file = not AUDIT_LOG.exists()
            prev = audit_last_hash()
            row = [str(date.today()), datetime.now().strftime("%H:%M:%S"),
                   str(uid), str(name),
                   (reason if action == "mark" and reason in
                    ("present", "late", "half_day")
                    else ("late" if late else "ontime")) if action == "mark"
                   else action,
                   str(xp), reason, actor]
            digest = hashlib.sha256(
                (prev + "|" + "|".join(row)).encode()).hexdigest()[:16]
            with open(AUDIT_LOG, "a", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                if new_file:
                    w.writerow(["date", "time", "uid", "name", "status",
                                "xp", "reason", "actor", "hash"])
                w.writerow(row + [digest])
        except Exception:
            pass

    def jpeg(self):
        with self.frame_lock:
            if self.frame is None:
                return None
            ok, buf = cv2.imencode(".jpg", self.frame, [int(cv2.IMWRITE_JPEG_QUALITY), 72])
            return buf.tobytes() if ok else None

    def grab_still(self):
        """One frame for enrolment photos — opens the camera if needed."""
        if cv2 is None:
            return None
        if self.running:
            with self.frame_lock:
                return self.frame.copy() if self.frame is not None else None
        cap = None
        for backend in self._backends():
            for idx in (0, 1, 2):
                try:
                    c = cv2.VideoCapture(idx, backend) if backend else cv2.VideoCapture(idx)
                except Exception:
                    continue
                if c.isOpened():
                    cap = c
                    break
                c.release()
            if cap:
                break
        if cap is None:
            return None
        for _ in range(8):        # let exposure settle
            cap.read()
            time.sleep(0.05)
        ok, frame = cap.read()
        cap.release()
        return frame if ok else None

# ══════════════════════════════════════════════════════════════════════════
#  EMBEDDED WEB ASSETS  (generated — edit web/ sources, then re-bundle)
# ══════════════════════════════════════════════════════════════════════════
APP_CSS = """/* ══════════════════════════════════════════════════════════════════
   FaceAttend 3.2 — stylesheet
   Dark, glassy, animated. Everything Tkinter could not do.
   ══════════════════════════════════════════════════════════════════ */

/* ── design tokens ─────────────────────────────────────────────── */
:root{
  --bg-0:#05070d;
  --bg-1:#090d16;
  --bg-2:#0e1420;

  --surface:rgba(255,255,255,.038);
  --surface-2:rgba(255,255,255,.06);
  --surface-3:rgba(255,255,255,.09);
  --stroke:rgba(255,255,255,.09);
  --stroke-2:rgba(255,255,255,.16);

  --text:#eef2fb;
  --muted:#94a3bb;
  --faint:#5b6880;

  --accent:#5b8cff;
  --accent-2:#a56bff;
  --accent-3:#22d3ee;
  --grad:linear-gradient(120deg,var(--accent) 0%,var(--accent-2) 55%,var(--accent-3) 100%);
  --grad-soft:linear-gradient(120deg,rgba(91,140,255,.16),rgba(165,107,255,.16));

  --ok:#2ee6a8;
  --warn:#ffbe3d;
  --bad:#ff5d73;
  --xp:#ffd166;

  --house-red:#ff5d73;
  --house-blue:#4d9dff;
  --house-green:#2ee6a8;
  --house-orange:#ffa23d;

  --r-sm:10px; --r:16px; --r-lg:22px; --r-xl:30px;
  --shadow:0 10px 40px -12px rgba(0,0,0,.7);
  --shadow-lg:0 28px 70px -22px rgba(0,0,0,.85);
  --glow:0 0 0 1px rgba(91,140,255,.35),0 12px 44px -10px rgba(91,140,255,.45);

  --ease:cubic-bezier(.22,1,.36,1);
  --rail:82px;
  --panel:340px;
  --font:'Inter','Ubuntu','Segoe UI',system-ui,-apple-system,sans-serif;
  --mono:'JetBrains Mono','Ubuntu Mono','DejaVu Sans Mono',monospace;
}

/* light theme override */
body.light{
  --bg-0:#eef1f7; --bg-1:#f6f8fc; --bg-2:#ffffff;
  --surface:rgba(15,23,42,.035); --surface-2:rgba(15,23,42,.055);
  --surface-3:rgba(15,23,42,.08);
  --stroke:rgba(15,23,42,.10); --stroke-2:rgba(15,23,42,.18);
  --text:#0d1526; --muted:#5a6a85; --faint:#8c9ab2;
  --shadow:0 10px 34px -14px rgba(15,23,42,.3);
  --shadow-lg:0 24px 60px -24px rgba(15,23,42,.35);
}

*{box-sizing:border-box;margin:0;padding:0}
::selection{background:rgba(91,140,255,.34);color:#fff}

html,body{height:100%}
body{
  font-family:var(--font);
  background:var(--bg-0);
  color:var(--text);
  overflow:hidden;
  -webkit-font-smoothing:antialiased;
  font-feature-settings:'cv02','cv03','ss01';
}

/* ── animated aurora backdrop ──────────────────────────────────── */
.aurora{position:fixed;inset:0;z-index:0;overflow:hidden;pointer-events:none}
.aurora b{
  position:absolute;border-radius:50%;filter:blur(90px);opacity:.30;
  animation:drift 26s var(--ease) infinite alternate;
}
.aurora b:nth-child(1){width:46vw;height:46vw;background:#2f6bff;top:-12vw;left:-8vw}
.aurora b:nth-child(2){width:38vw;height:38vw;background:#a56bff;bottom:-14vw;right:-6vw;animation-delay:-9s}
.aurora b:nth-child(3){width:30vw;height:30vw;background:#22d3ee;top:38%;left:52%;animation-delay:-17s;opacity:.16}
@keyframes drift{
  from{transform:translate3d(0,0,0) scale(1)}
  to{transform:translate3d(4vw,5vh,0) scale(1.22)}
}
body.light .aurora b{opacity:.18}
body.reduce-motion .aurora b{animation:none}

/* ── shell ─────────────────────────────────────────────────────── */
.shell{position:relative;z-index:1;display:grid;grid-template-columns:var(--rail) var(--panel) 1fr;height:100vh}
.shell.wide{grid-template-columns:var(--rail) 1fr}
.shell.wide .panel{display:none}

/* ── nav rail ──────────────────────────────────────────────────── */
.rail{
  display:flex;flex-direction:column;align-items:center;gap:6px;
  padding:18px 0;background:rgba(5,7,13,.55);
  backdrop-filter:blur(22px);-webkit-backdrop-filter:blur(22px);
  border-right:1px solid var(--stroke);
}
body.light .rail{background:rgba(255,255,255,.6)}

.brand{
  width:46px;height:46px;border-radius:14px;display:grid;place-items:center;
  background:var(--grad);color:#fff;font-weight:800;font-size:19px;
  box-shadow:var(--glow);margin-bottom:16px;cursor:pointer;
  transition:transform .5s var(--ease);
}
.brand:hover{transform:rotate(8deg) scale(1.08)}
.brand:active{transform:scale(.94)}

.navbtn{
  position:relative;width:50px;height:46px;border:0;border-radius:13px;
  background:transparent;color:var(--faint);font-size:19px;cursor:pointer;
  display:grid;place-items:center;transition:all .28s var(--ease);
}
.navbtn:hover{background:var(--surface-2);color:var(--text);transform:translateY(-2px)}
.navbtn.on{background:var(--surface-3);color:var(--text)}
.navbtn.on::before{
  content:'';position:absolute;left:-14px;top:50%;transform:translateY(-50%);
  width:3px;height:22px;border-radius:3px;background:var(--grad);
  box-shadow:0 0 14px var(--accent);
}
.navbtn .tip{
  position:absolute;left:60px;top:50%;transform:translateY(-50%) translateX(-6px);
  background:var(--bg-2);color:var(--text);font-size:12px;font-weight:600;
  padding:7px 11px;border-radius:9px;white-space:nowrap;opacity:0;
  pointer-events:none;transition:all .22s var(--ease);
  border:1px solid var(--stroke);box-shadow:var(--shadow);z-index:60;
}
.navbtn:hover .tip{opacity:1;transform:translateY(-50%) translateX(0)}

.rail-foot{margin-top:auto;display:flex;flex-direction:column;align-items:center;gap:10px}
.dot{width:9px;height:9px;border-radius:50%;background:var(--faint);transition:.3s}
.dot.live{background:var(--ok);box-shadow:0 0 0 0 rgba(46,230,168,.65);animation:pulse 1.9s infinite}
@keyframes pulse{70%{box-shadow:0 0 0 11px rgba(46,230,168,0)}100%{box-shadow:0 0 0 0 rgba(46,230,168,0)}}

/* ── side panel ────────────────────────────────────────────────── */
.panel{
  background:rgba(9,13,22,.5);backdrop-filter:blur(22px);
  -webkit-backdrop-filter:blur(22px);border-right:1px solid var(--stroke);
  overflow-y:auto;padding:22px;
}
body.light .panel{background:rgba(255,255,255,.55)}
.page{display:none;animation:fade .45s var(--ease)}
.page.on{display:block}
@keyframes fade{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}

.ptitle{font-size:19px;font-weight:800;letter-spacing:-.4px;display:flex;align-items:center;gap:9px}
.psub{color:var(--muted);font-size:12.5px;margin:5px 0 18px;line-height:1.55}

/* ── main ──────────────────────────────────────────────────────── */
.main{overflow-y:auto;padding:26px 30px 40px}
.head{display:flex;align-items:flex-end;justify-content:space-between;gap:18px;margin-bottom:22px;flex-wrap:wrap}
h1{
  font-size:33px;font-weight:850;letter-spacing:-1.1px;line-height:1.1;
  background:linear-gradient(180deg,var(--text),color-mix(in srgb,var(--text) 62%,transparent));
  -webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;
}
.date{color:var(--muted);font-size:13px;margin-top:5px;display:flex;align-items:center;gap:9px}
.clock{font-variant-numeric:tabular-nums;font-weight:700;color:var(--text)}

/* ── buttons ───────────────────────────────────────────────────── */
.btn{
  border:1px solid var(--stroke);background:var(--surface-2);color:var(--text);
  padding:10px 16px;border-radius:12px;font-size:13px;font-weight:650;
  cursor:pointer;font-family:inherit;display:inline-flex;align-items:center;
  gap:7px;transition:all .24s var(--ease);position:relative;overflow:hidden;
}
.btn:hover{background:var(--surface-3);transform:translateY(-2px);box-shadow:var(--shadow)}
.btn:active{transform:translateY(0) scale(.97)}
.btn:disabled{opacity:.42;cursor:not-allowed;transform:none}
.btn.primary{background:var(--grad);border-color:transparent;color:#fff;box-shadow:var(--glow)}
.btn.primary:hover{filter:brightness(1.12)}
.btn.ok{background:linear-gradient(120deg,#16a97a,#2ee6a8);border-color:transparent;color:#04241a}
.btn.warn{background:linear-gradient(120deg,#d98a00,#ffbe3d);border-color:transparent;color:#2a1c00}
.btn.bad{background:linear-gradient(120deg,#d63a52,#ff5d73);border-color:transparent;color:#fff}
.btn.sm{padding:7px 12px;font-size:12px;border-radius:10px}
.btn.block{width:100%;justify-content:center}
.btn.ghost{background:transparent}

/* ripple */
.btn .ripple{position:absolute;border-radius:50%;background:rgba(255,255,255,.42);
  transform:scale(0);animation:rip .62s var(--ease);pointer-events:none}
@keyframes rip{to{transform:scale(3.4);opacity:0}}

.row{display:flex;gap:10px;flex-wrap:wrap}

/* ── cards ─────────────────────────────────────────────────────── */
.card{
  background:var(--surface);border:1px solid var(--stroke);border-radius:var(--r-lg);
  padding:18px;backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);
  transition:all .34s var(--ease);
}
.card:hover{border-color:var(--stroke-2);box-shadow:var(--shadow)}

/* KPI strip */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(196px,1fr));gap:14px;margin-bottom:22px}
.kpi{
  position:relative;overflow:hidden;border-radius:var(--r-lg);padding:18px 18px 16px;
  background:var(--surface);border:1px solid var(--stroke);
  backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);
  transition:all .38s var(--ease);
}
.kpi::after{
  content:'';position:absolute;inset:0;border-radius:inherit;padding:1px;
  background:var(--grad);-webkit-mask:linear-gradient(#000 0 0) content-box,linear-gradient(#000 0 0);
  -webkit-mask-composite:xor;mask-composite:exclude;opacity:0;transition:opacity .38s var(--ease);
}
.kpi:hover{transform:translateY(-5px);box-shadow:var(--shadow-lg)}
.kpi:hover::after{opacity:.85}
.kpi .ic{
  width:38px;height:38px;border-radius:12px;display:grid;place-items:center;
  font-size:17px;margin-bottom:13px;
}
.kpi .lb{color:var(--muted);font-size:11.5px;font-weight:700;text-transform:uppercase;letter-spacing:.9px}
.kpi .vl{font-size:35px;font-weight:850;letter-spacing:-1.5px;margin:5px 0 3px;font-variant-numeric:tabular-nums}
.kpi .ft{color:var(--faint);font-size:11.5px}
.kpi .spark{position:absolute;right:0;bottom:0;width:96px;height:38px;opacity:.5}

.i-blue{background:rgba(91,140,255,.16);color:#8fb2ff}
.i-green{background:rgba(46,230,168,.16);color:var(--ok)}
.i-amber{background:rgba(255,190,61,.16);color:var(--warn)}
.i-purple{background:rgba(165,107,255,.16);color:#c79bff}
.i-red{background:rgba(255,93,115,.16);color:var(--bad)}

/* ── toolbar ───────────────────────────────────────────────────── */
.bar{display:flex;gap:9px;flex-wrap:wrap;align-items:center;margin-bottom:18px}
.search{
  flex:1;min-width:220px;position:relative;
}
.search input{
  width:100%;padding:11px 15px 11px 40px;border-radius:13px;
  border:1px solid var(--stroke);background:var(--surface);color:var(--text);
  font-family:inherit;font-size:13.5px;outline:none;transition:all .26s var(--ease);
}
.search input:focus{border-color:var(--accent);box-shadow:0 0 0 4px rgba(91,140,255,.13);background:var(--surface-2)}
.search svg{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--faint);pointer-events:none}
.search kbd{position:absolute;right:11px;top:50%;transform:translateY(-50%)}

kbd{
  font-family:var(--mono);font-size:10.5px;padding:3px 6px;border-radius:6px;
  background:var(--surface-2);border:1px solid var(--stroke);color:var(--muted);
}

/* ── table ─────────────────────────────────────────────────────── */
.tablewrap{
  border-radius:var(--r-lg);border:1px solid var(--stroke);overflow:hidden;
  background:var(--surface);backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);
}
.scroll{overflow:auto;max-height:calc(100vh - 400px)}
table{width:100%;border-collapse:collapse;font-size:13.5px}
thead th{
  position:sticky;top:0;z-index:5;background:var(--bg-2);
  padding:13px 15px;text-align:left;font-size:11px;font-weight:750;
  text-transform:uppercase;letter-spacing:.9px;color:var(--muted);
  border-bottom:1px solid var(--stroke);white-space:nowrap;cursor:pointer;
  user-select:none;transition:color .2s;
}
thead th:hover{color:var(--text)}
thead th .ar{opacity:.4;font-size:9px;margin-left:3px}
tbody tr{border-bottom:1px solid var(--stroke);transition:background .2s}
tbody tr:hover{background:var(--surface-2)}
tbody tr.sel{background:rgba(91,140,255,.1)}
tbody tr.flash{animation:flashrow 1.5s var(--ease)}
@keyframes flashrow{
  0%{background:rgba(46,230,168,.3)}100%{background:transparent}
}
tbody td{padding:11px 15px;vertical-align:middle;white-space:nowrap}

.who{display:flex;align-items:center;gap:11px;min-width:180px}
.av{
  width:38px;height:38px;border-radius:11px;object-fit:cover;flex:none;
  border:1.5px solid var(--stroke-2);background:var(--surface-2);
}
.av.ph{display:grid;place-items:center;font-weight:750;font-size:14px;color:#fff}
.who b{font-weight:650;font-size:13.5px;display:block}
.who span{color:var(--faint);font-size:11.5px}

.pill{
  display:inline-flex;align-items:center;gap:5px;padding:4px 10px;
  border-radius:99px;font-size:11.5px;font-weight:700;white-space:nowrap;
}
.pill.on{background:rgba(46,230,168,.14);color:var(--ok)}
.pill.off{background:rgba(255,93,115,.13);color:var(--bad)}
.pill.late{background:rgba(255,190,61,.14);color:var(--warn)}
.pill.xp{background:rgba(255,209,102,.13);color:var(--xp)}

.hd{width:9px;height:9px;border-radius:50%;display:inline-block;margin-right:7px}

.empty{padding:64px 24px;text-align:center;color:var(--faint)}
.empty .em{font-size:46px;margin-bottom:14px;opacity:.5}
.empty h3{font-size:16px;color:var(--muted);margin-bottom:7px;font-weight:700}

/* ── camera ────────────────────────────────────────────────────── */
.cam{
  position:relative;width:100%;aspect-ratio:4/3;border-radius:var(--r-lg);
  overflow:hidden;background:#04070d;border:1px solid var(--stroke);
  display:grid;place-items:center;
}
.cam img{width:100%;height:100%;object-fit:cover;display:block}
.cam .off{text-align:center;color:var(--faint);padding:20px}
.cam .off .ic{font-size:38px;margin-bottom:9px;opacity:.55}
.cam .badge{
  position:absolute;top:11px;left:11px;display:flex;align-items:center;gap:6px;
  background:rgba(5,7,13,.72);backdrop-filter:blur(10px);padding:6px 11px;
  border-radius:99px;font-size:11px;font-weight:750;border:1px solid var(--stroke);
}
.cam .badge .dot{width:7px;height:7px}
@keyframes scan{0%,100%{top:6%;opacity:.35}50%{top:94%;opacity:1}}
body.reduce-motion 
/* recognition card */
.recog{
  margin-top:14px;padding:16px;border-radius:var(--r);
  background:var(--grad-soft);border:1px solid var(--stroke-2);
  transition:all .4s var(--ease);
}
.recog.hit{animation:hit .8s var(--ease)}
@keyframes hit{
  0%{transform:scale(1)}30%{transform:scale(1.04);box-shadow:var(--glow)}100%{transform:scale(1)}
}
.recog .lb{font-size:10.5px;text-transform:uppercase;letter-spacing:1.1px;color:var(--muted);font-weight:750}
.recog h3{font-size:19px;font-weight:800;margin:6px 0 5px;letter-spacing:-.5px}
.recog .meta{font-size:12px;color:var(--muted);display:flex;gap:8px;flex-wrap:wrap;align-items:center}

/* ── feed / console ────────────────────────────────────────────── */
.feed{
  margin-top:10px;max-height:230px;min-height:56px;overflow-y:auto;
  border-radius:var(--r);border:1px solid var(--stroke);
  background:rgba(0,0,0,.22);padding:6px;
}
/* Empty feed used to collapse to a bare line that overlapped the heading. */
.feed:empty::after{
  content:'Nothing yet — start a scan and recognitions appear here.';
  display:block;padding:16px 12px;text-align:center;
  font-size:11.5px;color:var(--faint);line-height:1.5;
}
.fi{
  display:flex;gap:9px;padding:8px 10px;border-radius:9px;font-size:12.5px;
  animation:slide .4s var(--ease);
}
.fi:hover{background:var(--surface-2)}
.fi time{color:var(--faint);font-family:var(--mono);font-size:11px;flex:none}
@keyframes slide{from{opacity:0;transform:translateX(-10px)}to{opacity:1;transform:none}}

.console{
  font-family:var(--mono);font-size:12px;line-height:1.75;
  background:#05080f;border-radius:0;padding:16px 18px;
  min-height:0;overflow-y:auto;border:0;flex:1;
}
.console .ln{display:flex;gap:11px;padding:2px 0;min-height:21px;animation:slide .22s var(--ease)}
.console .t{color:#56647d;flex:none;user-select:none}
.console .ln::before{content:'›';color:#34445f;flex:none}
.console .cmd{color:#d7e2f3}
.console .i{color:var(--accent-3)}
.console .s{color:var(--ok)}
.console .w{color:var(--warn)}
.console .e{color:var(--bad)}

/* ── full-width 3.2 panes ─────────────────────────────────────── */
.page.wide-page{
  min-height:100%;padding:4px 2px 24px;
}
.page.wide-page > .ptitle{font-size:25px;margin-top:2px}
.wide-head{
  display:flex;align-items:flex-end;justify-content:space-between;gap:18px;
  margin:2px 0 20px;flex-wrap:wrap;
}
.wide-head .ptitle{font-size:25px}
.pane-switch{display:flex;align-items:center;gap:8px;margin:0 0 18px}
.pane-switch .pane-label{font-size:12px;color:var(--muted);font-weight:700;margin-right:4px}
.pane-btn{
  border:1px solid var(--stroke);background:var(--surface-2);color:var(--muted);
  padding:8px 13px;border-radius:10px;font:650 12px var(--font);cursor:pointer;
  transition:all .22s var(--ease)
}
.pane-btn:hover{background:var(--surface-3);color:var(--text)}
.pane-btn.on{background:var(--accent);border-color:transparent;color:#fff;box-shadow:0 5px 18px -8px var(--accent)}
.pane-body{animation:fade .3s var(--ease)}
.page[data-page="analytics"].wide-page .analytics-grid{
  display:grid;grid-template-columns:minmax(0,1.25fr) minmax(300px,.75fr);gap:16px
}
.page[data-page="analytics"].wide-page .analytics-main,
.page[data-page="analytics"].wide-page .analytics-side{min-width:0}
.page[data-page="analytics"].wide-page .card{height:100%}

/* 3.2 notepad: saved notes beside a real white page */
.notepad-layout{display:grid;grid-template-columns:220px minmax(0,1fr);gap:14px;height:calc(100vh - 174px);min-height:420px}
.note-sidebar,.note-editor{
  border:1px solid var(--stroke);background:var(--surface);border-radius:var(--r-lg);
  min-height:0;overflow:hidden
}
.note-sidebar{padding:14px;display:flex;flex-direction:column}
.note-sidebar .btn{margin-bottom:8px}
.note-sidebar .sec{margin:14px 0 8px}
.note-sidebar .notelist{max-height:none;overflow-y:auto;flex:1}
.note-editor{background:#d9d9d9;display:grid;grid-template-rows:auto auto minmax(0,1fr);gap:0}
.note-editor-top,.note-editor-tools{display:flex;align-items:center;gap:8px;padding:8px 10px;background:var(--surface)}
.note-editor-tools{background:#ededed;padding:5px 8px}
.note-editor-top [contenteditable]{flex:1;min-width:0;color:var(--text);font-size:15px;font-weight:750;outline:none}
.note-editor-tools .tool{
  min-width:28px;height:27px;padding:0 8px;border:1px solid #c7c7c7;border-radius:6px;
  background:#fff;color:#111;cursor:pointer;font:700 12px var(--mono)
}
.note-editor-tools .tool:hover{background:#e5e7eb}
.note-page{min-height:0;margin:0 1px 1px;background:#fff;display:flex}
#noteBody{
  width:100%;height:100%;padding:18px 20px;border:0;border-radius:0;background:#fff;
  color:#111;font-family:var(--mono);font-size:13px;line-height:1.7;outline:none;resize:none
}
#noteBody:focus{box-shadow:inset 0 0 0 2px rgba(91,140,255,.35)}
.whiteboard-layout{display:grid;grid-template-columns:220px minmax(0,1fr);gap:14px;min-height:calc(100vh - 174px)}
.whiteboard-tools{border:1px solid var(--stroke);border-radius:var(--r-lg);background:var(--surface);padding:16px}
.whiteboard-stage{min-width:0;border:1px solid var(--stroke);border-radius:var(--r-lg);background:#d9d9d9;padding:10px;overflow:auto}
.whiteboard-stage #board{width:100%;height:auto;min-height:420px;object-fit:contain}

/* a terminal, not a text tab */
.console-layout{display:grid;grid-template-columns:minmax(0,1fr) 250px;gap:15px;min-height:calc(100vh - 176px)}
.terminal-shell{display:flex;flex-direction:column;min-height:520px;overflow:hidden;border:1px solid #263653;border-radius:var(--r-lg);background:#05080f;box-shadow:0 18px 45px -25px rgba(0,0,0,.9)}
.terminal-bar{display:flex;align-items:center;gap:9px;padding:11px 14px;border-bottom:1px solid #1b2a43;background:#0b1220;color:#8fa4c2;font:700 11px var(--mono);letter-spacing:.35px}
.terminal-lights{display:flex;gap:5px;margin-right:4px}
.terminal-lights i{width:8px;height:8px;border-radius:50%;display:block}
.terminal-lights i:nth-child(1){background:#ff5d73}.terminal-lights i:nth-child(2){background:#ffbe3d}.terminal-lights i:nth-child(3){background:#2ee6a8}
.terminal-live{margin-left:auto;color:var(--ok);font-size:10px}
.terminal-live::before{content:'';display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--ok);margin:0 5px 1px 0;box-shadow:0 0 8px var(--ok)}
.console-prompt{display:flex;align-items:center;gap:10px;padding:10px 14px;border-top:1px solid #1b2a43;background:#0a111e}
.console-prompt .prompt{font:700 12px var(--mono);color:var(--accent-3);white-space:nowrap}
.console-prompt input{flex:1;min-width:0;border:0;outline:none;background:transparent;color:#e7edf7;font:12px var(--mono)}
.console-prompt input::placeholder{color:#53627b}
.console-prompt button{border:0;border-radius:7px;background:var(--accent);color:#fff;padding:7px 12px;font:700 11px var(--font);cursor:pointer}
.console-prompt button:hover{filter:brightness(1.14)}
.console-inspector{display:flex;flex-direction:column;gap:12px}
.console-inspector .card{padding:15px}
.console-inspector .sec{margin:3px 0 0}
.command-list{font:11px/1.9 var(--mono);color:var(--muted)}
.command-list b{color:var(--accent-3);font-weight:600}

/* ── forms ─────────────────────────────────────────────────────── */
.field{margin-bottom:14px}
.field label{display:block;font-size:11.5px;font-weight:700;color:var(--muted);margin-bottom:6px;text-transform:uppercase;letter-spacing:.6px}
.field input,.field select,.field textarea{
  width:100%;padding:11px 13px;border-radius:11px;border:1px solid var(--stroke);
  background:var(--surface);color:var(--text);font-family:inherit;font-size:13.5px;
  outline:none;transition:all .24s var(--ease);
}
.field input:focus,.field select:focus,.field textarea:focus{
  border-color:var(--accent);box-shadow:0 0 0 4px rgba(91,140,255,.13);background:var(--surface-2)
}
.field textarea{resize:vertical;min-height:90px;font-size:13px}
.field select option{background:var(--bg-2);color:var(--text)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:11px}

/* toggle */
.tg{display:flex;align-items:center;justify-content:space-between;padding:11px 0;border-bottom:1px solid var(--stroke)}
.tg:last-child{border:0}
.tg .tl{font-size:13px;font-weight:600}
.tg .td{font-size:11.5px;color:var(--faint);margin-top:2px}
.sw{
  position:relative;width:44px;height:25px;border-radius:99px;flex:none;
  background:var(--surface-3);border:1px solid var(--stroke);cursor:pointer;transition:.3s var(--ease);
}
.sw::after{
  content:'';position:absolute;top:2.5px;left:3px;width:18px;height:18px;
  border-radius:50%;background:var(--muted);transition:.3s var(--ease);
}
.sw.on{background:var(--grad);border-color:transparent}
.sw.on::after{transform:translateX(19px);background:#fff}

/* ── modal ─────────────────────────────────────────────────────── */
.veil{
  position:fixed;inset:0;z-index:100;background:rgba(3,5,10,.72);
  backdrop-filter:blur(9px);-webkit-backdrop-filter:blur(9px);
  display:none;place-items:center;padding:20px;
}
.veil.on{display:grid;animation:fade .28s var(--ease)}
.modal{
  width:100%;max-width:560px;max-height:88vh;overflow-y:auto;
  background:var(--bg-1);border:1px solid var(--stroke-2);border-radius:var(--r-xl);
  padding:26px;box-shadow:var(--shadow-lg);animation:pop .42s var(--ease);
}
@keyframes pop{from{opacity:0;transform:scale(.93) translateY(22px)}to{opacity:1;transform:none}}
.modal h2{font-size:21px;font-weight:800;margin-bottom:5px;letter-spacing:-.5px}
.modal .sub{color:var(--muted);font-size:13px;margin-bottom:20px}
.modal .acts{display:flex;gap:10px;justify-content:flex-end;margin-top:22px}

/* ── toasts ────────────────────────────────────────────────────── */
.toasts{position:fixed;bottom:24px;right:24px;z-index:200;display:flex;flex-direction:column;gap:10px}
.toast{
  min-width:270px;max-width:400px;padding:14px 16px;border-radius:15px;
  background:var(--bg-2);border:1px solid var(--stroke-2);box-shadow:var(--shadow-lg);
  display:flex;gap:12px;align-items:flex-start;animation:tin .45s var(--ease);
  backdrop-filter:blur(18px);
}
.toast.out{animation:tout .35s var(--ease) forwards}
@keyframes tin{from{opacity:0;transform:translateX(90px) scale(.94)}to{opacity:1;transform:none}}
@keyframes tout{to{opacity:0;transform:translateX(90px) scale(.94)}}
.toast .ti{width:30px;height:30px;border-radius:9px;display:grid;place-items:center;flex:none;font-size:14px}
.toast b{display:block;font-size:13.5px;font-weight:700;margin-bottom:2px}
.toast p{font-size:12.5px;color:var(--muted);line-height:1.5}
.toast .bar{position:absolute;bottom:0;left:0;height:2px;background:var(--grad);animation:tbar 4s linear}
@keyframes tbar{from{width:100%}to{width:0}}

/* ── charts ────────────────────────────────────────────────────── */
.chart{width:100%;height:150px}
.bars{display:flex;align-items:flex-end;gap:9px;height:150px;padding-top:22px}
.bcol{flex:1;height:100%;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;gap:7px}
.bcol .bk{
  width:100%;border-radius:8px 8px 3px 3px;min-height:4px;
  transition:height 1.1s var(--ease);position:relative;
}
.bcol .bk::after{
  content:attr(data-v);position:absolute;top:-19px;left:50%;transform:translateX(-50%);
  font-size:10.5px;font-weight:750;color:var(--muted);white-space:nowrap;
}
.bcol .bl{font-size:11px;color:var(--faint);font-weight:600}

/* house cards */
.houses{display:grid;grid-template-columns:1fr 1fr;gap:11px}
.hcard{
  padding:15px;border-radius:var(--r);border:1px solid var(--stroke);
  position:relative;overflow:hidden;transition:all .34s var(--ease);
}
.hcard:hover{transform:translateY(-3px);box-shadow:var(--shadow)}
.hcard .hn{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:750}
.hcard .hp{font-size:29px;font-weight:850;letter-spacing:-1.2px;margin:7px 0 2px;font-variant-numeric:tabular-nums}
.hcard .hs{font-size:11.5px;color:var(--muted)}
.hcard .hbar{position:absolute;left:0;bottom:0;height:3px;transition:width 1.1s var(--ease)}
.hcard.lead::after{content:'👑';position:absolute;top:11px;right:13px;font-size:16px}

/* ══ AI BUDDY ═══════════════════════════════════════════════════ */
.bothead{
  position:relative;overflow:hidden;border-radius:var(--r-lg);padding:16px 17px;
  background:linear-gradient(135deg,rgba(91,140,255,.20),rgba(165,107,255,.14) 55%,rgba(34,211,238,.10));
  border:1px solid var(--stroke-2);margin-bottom:14px;
}
.bothead::before{
  content:'';position:absolute;width:180px;height:180px;border-radius:50%;
  background:radial-gradient(circle,rgba(165,107,255,.30),transparent 70%);
  top:-90px;right:-60px;filter:blur(28px);animation:drift 15s var(--ease) infinite alternate;
}
body.reduce-motion .bothead::before{animation:none}
.bothead .r1{display:flex;align-items:center;gap:12px;position:relative;z-index:1}
.avatar{
  width:44px;height:44px;border-radius:14px;flex:none;display:grid;place-items:center;
  font-size:21px;background:var(--grad);box-shadow:var(--glow);
  animation:bob 4.5s var(--ease) infinite;
}
@keyframes bob{0%,100%{transform:translateY(0) rotate(0)}50%{transform:translateY(-4px) rotate(-4deg)}}
body.reduce-motion .avatar{animation:none}
.bothead h3{font-size:17px;font-weight:800;letter-spacing:-.4px;line-height:1.2}
.bothead .status{display:flex;align-items:center;gap:6px;font-size:11.5px;color:var(--muted);margin-top:3px}
.bothead .status .d{width:7px;height:7px;border-radius:50%;background:var(--faint);flex:none}
.bothead .status.ai .d{background:var(--ok);box-shadow:0 0 9px var(--ok);animation:pulse 2s infinite}
.bothead .status.ai b{color:var(--ok)}
.bothead p{font-size:12px;color:var(--muted);line-height:1.55;margin-top:11px;position:relative;z-index:1}

.chat{
  flex:1;min-height:180px;overflow-y:auto;border-radius:var(--r-lg);
  border:1px solid var(--stroke);background:rgba(0,0,0,.26);
  padding:15px;display:flex;flex-direction:column;gap:14px;
}
.turn{display:flex;gap:10px;animation:slide .38s var(--ease)}
.turn.you{flex-direction:row-reverse}
.turn .pic{
  width:29px;height:29px;border-radius:9px;flex:none;display:grid;place-items:center;
  font-size:14px;margin-top:2px;
}
.turn.bot .pic{background:var(--grad);box-shadow:0 3px 12px -3px rgba(91,140,255,.6)}
.turn.you .pic{background:var(--surface-3);border:1px solid var(--stroke-2);font-size:12px;font-weight:700;color:var(--muted)}
.turn .body{max-width:calc(100% - 44px);display:flex;flex-direction:column;gap:3px}
.turn.you .body{align-items:flex-end}
.msg{
  padding:11px 14px;border-radius:16px;font-size:13px;line-height:1.65;
  white-space:pre-wrap;word-break:break-word;
}
.turn.bot .msg{background:var(--surface-2);border:1px solid var(--stroke);border-top-left-radius:5px}
.turn.you .msg{background:var(--grad);color:#fff;border-top-right-radius:5px;box-shadow:0 4px 16px -6px rgba(91,140,255,.7)}
.turn .when{font-size:10px;color:var(--faint);padding:0 5px;font-variant-numeric:tabular-nums}
.msg.think{display:flex;gap:5px;padding:14px 16px;align-items:center}
.msg.think i{width:6px;height:6px;border-radius:50%;background:var(--muted);animation:blink 1.4s infinite}
.msg.think i:nth-child(2){animation-delay:.18s}
.msg.think i:nth-child(3){animation-delay:.36s}
@keyframes blink{0%,60%,100%{opacity:.28;transform:translateY(0)}30%{opacity:1;transform:translateY(-5px)}}

.chiprow{display:flex;gap:7px;flex-wrap:wrap;margin-top:12px}
.chip{
  padding:7px 12px;border-radius:99px;border:1px solid var(--stroke);
  background:var(--surface);color:var(--muted);font-size:12px;font-weight:650;
  cursor:pointer;transition:all .26s var(--ease);font-family:inherit;
  display:inline-flex;align-items:center;gap:5px;
}
.chip:hover{
  background:var(--grad-soft);color:var(--text);transform:translateY(-2px);
  border-color:var(--accent);box-shadow:0 5px 16px -8px var(--accent);
}
.chip:active{transform:translateY(0) scale(.96)}

.ask{
  display:flex;gap:8px;margin-top:12px;padding:6px;border-radius:16px;
  background:var(--surface);border:1px solid var(--stroke);
  transition:all .26s var(--ease);
}
.ask:focus-within{border-color:var(--accent);box-shadow:0 0 0 4px rgba(91,140,255,.13)}
.ask input{
  flex:1;padding:9px 12px;border:0;background:transparent;color:var(--text);
  font-family:inherit;font-size:13.5px;outline:none;
}
.ask button{
  width:38px;height:38px;border:0;border-radius:11px;background:var(--grad);
  color:#fff;font-size:15px;cursor:pointer;flex:none;
  display:grid;place-items:center;transition:all .24s var(--ease);
}
.ask button:hover{filter:brightness(1.15);transform:scale(1.06)}
.ask button:active{transform:scale(.93)}
.ask button:disabled{opacity:.4;cursor:not-allowed;transform:none}

/* the bot page becomes a full-height flex column */
.page[data-page="bot"].on{display:flex;flex-direction:column;height:100%}

/* ── whiteboard / notes ────────────────────────────────────────── */
#board{border-radius:var(--r-lg);border:1px solid var(--stroke);background:#fff;cursor:crosshair;display:block;max-width:100%;touch-action:none}
.swatches{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:13px}
.sw2{width:30px;height:30px;border-radius:9px;border:2px solid transparent;cursor:pointer;transition:.24s var(--ease)}
.sw2:hover{transform:scale(1.14)}
.sw2.on{border-color:var(--text);transform:scale(1.14)}
.notelist{display:flex;flex-direction:column;gap:7px;margin-top:11px;max-height:300px;overflow-y:auto}
.noteitem{padding:10px 12px;border-radius:11px;background:var(--surface);border:1px solid var(--stroke);cursor:pointer;font-size:13px;transition:.22s;display:flex;justify-content:space-between;align-items:center;gap:8px}
.noteitem:hover{background:var(--surface-3)}
.noteitem.on{border-color:var(--accent);background:rgba(91,140,255,.1)}
.noteitem small{color:var(--faint);font-size:11px}
#noteBody{width:100%;height:100%;padding:18px 20px;border:0;border-radius:0;background:#fff;color:#111;font-family:var(--mono);font-size:13px;line-height:1.7;outline:none;resize:none}
#noteBody:focus{box-shadow:inset 0 0 0 2px rgba(91,140,255,.35)}

/* leaderboard */
.lb{display:flex;flex-direction:column;gap:8px}
.lbrow{display:flex;align-items:center;gap:12px;padding:11px 13px;border-radius:13px;background:var(--surface);border:1px solid var(--stroke);transition:.3s var(--ease)}
.lbrow:hover{transform:translateX(4px);border-color:var(--stroke-2)}
.lbrow .rk{width:26px;height:26px;border-radius:8px;display:grid;place-items:center;font-size:12px;font-weight:800;background:var(--surface-2);color:var(--muted);flex:none}
.lbrow.p1 .rk{background:linear-gradient(135deg,#ffd166,#ff9f1c);color:#3d2600}
.lbrow.p2 .rk{background:linear-gradient(135deg,#d8dee9,#9aa5b8);color:#1e222b}
.lbrow.p3 .rk{background:linear-gradient(135deg,#cd7f32,#a86428);color:#fff}
.lbrow .nm{flex:1;font-size:13px;font-weight:650}
.lbrow .nm small{display:block;color:var(--faint);font-weight:500;font-size:11px}

/* misc */
.sec{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:1.1px;
  color:var(--faint);margin:20px 0 8px;line-height:1.4}
.sec:first-child{margin-top:0}
.spin{width:15px;height:15px;border:2px solid var(--stroke-2);border-top-color:var(--accent);border-radius:50%;animation:sp .7s linear infinite;display:inline-block}
@keyframes sp{to{transform:rotate(360deg)}}
.stat{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid var(--stroke);font-size:13px}
.stat:last-child{border:0}
.stat b{font-weight:700;font-variant-numeric:tabular-nums}

::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--surface-3);border-radius:9px}
::-webkit-scrollbar-thumb:hover{background:var(--stroke-2)}

/* confetti */
.conf{position:fixed;inset:0;pointer-events:none;z-index:300}
.cf{position:absolute;width:9px;height:9px;animation:fall linear forwards}
@keyframes fall{to{transform:translateY(105vh) rotate(760deg);opacity:0}}

/* party mode easter egg */
body.party .aurora b{animation-duration:2.6s!important}
body.party .brand{animation:spin 1.1s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
body.party h1{animation:hue 2.6s linear infinite}
@keyframes hue{to{filter:hue-rotate(360deg)}}

/* responsive */
@media(max-width:1180px){
  :root{--panel:300px}
  .kpi .vl{font-size:30px}
}
@media(max-width:900px){
  .shell{grid-template-columns:var(--rail) 1fr}
  .panel{display:none}
  .shell.showpanel{grid-template-columns:var(--rail) 1fr}
  .shell.showpanel .panel{display:block;position:fixed;left:var(--rail);top:0;bottom:0;width:320px;z-index:50}
  h1{font-size:26px}
  .main{padding:18px}
  .page.wide-page{padding:0 0 18px}
  .page[data-page="analytics"].wide-page .analytics-grid,
  .console-layout{grid-template-columns:1fr}
  .console-inspector{display:grid;grid-template-columns:1fr 1fr}
  .notepad-layout,.whiteboard-layout{grid-template-columns:1fr;height:auto;min-height:0}
  .note-sidebar{min-height:220px}
  .note-editor{height:calc(100vh - 380px);min-height:360px}
  .whiteboard-tools{min-height:0}
}

/* ══ SIGN-IN GATE ═══════════════════════════════════════════════ */
.gate{
  position:fixed;inset:0;z-index:400;display:grid;place-items:center;
  background:var(--bg-0);padding:24px;overflow-y:auto;
}
.gate.hidden{display:none}
.gate::before{
  content:'';position:absolute;width:60vw;height:60vw;border-radius:50%;
  background:radial-gradient(circle,rgba(91,140,255,.20),transparent 68%);
  top:-24vw;left:-14vw;filter:blur(70px);
  animation:drift 22s var(--ease) infinite alternate;
}
.gate::after{
  content:'';position:absolute;width:48vw;height:48vw;border-radius:50%;
  background:radial-gradient(circle,rgba(165,107,255,.18),transparent 68%);
  bottom:-20vw;right:-12vw;filter:blur(70px);
  animation:drift 26s var(--ease) infinite alternate-reverse;
}
body.reduce-motion .gate::before,body.reduce-motion .gate::after{animation:none}

.gatecard{
  position:relative;z-index:1;width:100%;max-width:400px;
  background:var(--surface);border:1px solid var(--stroke-2);
  border-radius:var(--r-xl);padding:32px 30px 24px;
  backdrop-filter:blur(22px);-webkit-backdrop-filter:blur(22px);
  box-shadow:var(--shadow-lg);animation:pop .5s var(--ease);
}
.gatebrand{display:flex;align-items:center;gap:14px;margin-bottom:26px}
.gatelogo{
  width:52px;height:52px;border-radius:16px;flex:none;display:grid;
  place-items:center;background:var(--grad);color:#fff;
  font-weight:850;font-size:19px;box-shadow:var(--glow);
}
.gatetitle{font-size:23px;font-weight:850;letter-spacing:-.7px;line-height:1.15}
.gatesub{font-size:12.5px;color:var(--muted);margin-top:3px}

.gatefield{margin-bottom:15px}
.gatefield label{
  display:block;font-size:11px;font-weight:750;color:var(--muted);
  margin-bottom:6px;text-transform:uppercase;letter-spacing:.7px;
}
.gatefield input{
  width:100%;padding:12px 14px;border-radius:12px;
  border:1px solid var(--stroke);background:var(--surface-2);
  color:var(--text);font-family:inherit;font-size:14px;outline:none;
  transition:all .24s var(--ease);
}
.gatefield input:focus{
  border-color:var(--accent);background:var(--surface-3);
  box-shadow:0 0 0 4px rgba(91,140,255,.14);
}
.gatepass{position:relative}
.gatepass input{padding-right:44px}
.gateeye{
  position:absolute;right:6px;top:50%;transform:translateY(-50%);
  width:32px;height:32px;border:0;border-radius:9px;background:transparent;
  color:var(--faint);cursor:pointer;font-size:14px;transition:.2s;
}
.gateeye:hover{background:var(--surface-3);color:var(--text)}

.gatemeter{height:3px;border-radius:2px;background:var(--stroke);margin-top:8px;overflow:hidden}
.gatemeter i{display:block;height:100%;width:0;border-radius:2px;transition:all .32s var(--ease)}

.gateerr{
  font-size:12.5px;color:var(--bad);min-height:18px;margin:2px 0 4px;
  line-height:1.45;
}
.gateswap{
  display:flex;align-items:center;justify-content:center;gap:6px;
  margin-top:18px;font-size:12.5px;color:var(--muted);
}
.gateswap button{
  border:0;background:transparent;color:var(--accent);
  font-family:inherit;font-size:12.5px;font-weight:700;cursor:pointer;
  padding:2px 4px;border-radius:6px;transition:.2s;
}
.gateswap button:hover{background:rgba(91,140,255,.12)}
.gatenote{
  margin-top:20px;padding-top:16px;border-top:1px solid var(--stroke);
  font-size:11px;color:var(--faint);line-height:1.6;text-align:center;
}
@media(max-width:460px){ .gatecard{padding:26px 22px 20px} }

/* ── high contrast ──────────────────────────────────────────────── */
body.contrast{
  --text:#ffffff; --muted:#d3dcea; --faint:#a9b6c8;
  --stroke:rgba(255,255,255,.34); --stroke-2:rgba(255,255,255,.52);
  --surface:rgba(255,255,255,.10); --surface-2:rgba(255,255,255,.15);
  --surface-3:rgba(255,255,255,.22);
  --accent:#7aa5ff; --ok:#4dffb8; --warn:#ffc94d; --bad:#ff7d8f;
}
body.contrast .aurora{opacity:.35}
body.contrast .kpi .lb,body.contrast .sec{color:#e6ecf6}
body.contrast .card,body.contrast .kpi{border-width:1.5px}
body.contrast .btn{border-width:1.5px;font-weight:750}
body.contrast .pill{font-weight:800}
body.contrast.light{
  --text:#000000; --muted:#22303f; --faint:#3d4b5c;
  --stroke:rgba(0,0,0,.42); --stroke-2:rgba(0,0,0,.6);
}

/* ── account picker ─────────────────────────────────────────────── */
.pickgrid{display:flex;flex-direction:column;gap:9px;margin-bottom:6px}
.picktile{
  display:flex;align-items:center;gap:13px;padding:11px 13px;
  border-radius:14px;background:var(--surface-2);border:1px solid var(--stroke);
  cursor:pointer;transition:all .24s var(--ease);text-align:left;
  font-family:inherit;width:100%;
}
.picktile:hover{
  background:var(--surface-3);border-color:var(--accent);
  transform:translateX(3px);
}
.picktile .pav{
  width:40px;height:40px;border-radius:13px;flex:none;display:grid;
  place-items:center;font-weight:800;font-size:15px;color:#fff;
}
.picktile .pinfo{flex:1;min-width:0}
.picktile .pname{
  font-size:13.5px;font-weight:700;color:var(--text);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
}
.picktile .prole{font-size:11px;color:var(--faint);margin-top:2px}
.picktile .parrow{color:var(--faint);font-size:15px}
.pickother,.pickback{
  width:100%;padding:10px;border:1px dashed var(--stroke-2);
  border-radius:12px;background:transparent;color:var(--muted);
  font-family:inherit;font-size:12.5px;font-weight:650;cursor:pointer;
  transition:all .22s var(--ease);margin-top:4px;
}
.pickother:hover,.pickback:hover{
  border-color:var(--accent);color:var(--text);background:var(--surface);
}
.pickback{margin:0 0 14px;text-align:left;padding:8px 10px;border-style:solid}
.gatewho{
  display:flex;align-items:center;gap:11px;padding:10px 12px;
  border-radius:13px;background:var(--surface-2);
  border:1px solid var(--stroke);margin-bottom:15px;
}
.gatewho .pav{width:36px;height:36px;border-radius:11px;display:grid;
  place-items:center;font-weight:800;font-size:14px;color:#fff;flex:none}

/* ══ ARCADE ═════════════════════════════════════════════════════ */
.arcade{
  position:fixed;inset:0;z-index:150;background:rgba(3,5,10,.80);
  backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);
  display:none;place-items:center;padding:20px;overflow-y:auto;
}
.arcade.on{display:grid;animation:fade .3s var(--ease)}
.arcadebox{
  width:100%;max-width:860px;background:var(--bg-1);
  border:1px solid var(--stroke-2);border-radius:var(--r-xl);
  padding:24px;box-shadow:var(--shadow-lg);animation:pop .45s var(--ease);
  max-height:94vh;overflow-y:auto;
}
.arcade-head{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;margin-bottom:20px}
.arcade-head h2{font-size:23px;font-weight:850;letter-spacing:-.6px}
.arcade-head p{font-size:12.5px;color:var(--muted);margin-top:4px;line-height:1.5}

.gamegrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(158px,1fr));gap:13px}
.gamecard{
  position:relative;overflow:hidden;padding:20px 16px;border-radius:var(--r-lg);
  background:var(--surface);border:1px solid var(--stroke);cursor:pointer;
  text-align:center;transition:all .34s var(--ease);
}
.gamecard::after{
  content:'';position:absolute;inset:0;border-radius:inherit;padding:1px;
  background:var(--grad);-webkit-mask:linear-gradient(#000 0 0) content-box,linear-gradient(#000 0 0);
  -webkit-mask-composite:xor;mask-composite:exclude;opacity:0;transition:opacity .34s var(--ease);
}
.gamecard:hover{transform:translateY(-5px);box-shadow:var(--shadow-lg);background:var(--surface-2)}
.gamecard:hover::after{opacity:.9}
.gamecard:active{transform:translateY(-1px) scale(.98)}
.gamecard .gi{font-size:34px;margin-bottom:10px;transition:transform .4s var(--ease)}
.gamecard:hover .gi{transform:scale(1.18) rotate(-7deg)}
.gamecard b{display:block;font-size:14.5px;font-weight:750;margin-bottom:5px}
.gamecard span{display:block;font-size:11.5px;color:var(--muted);line-height:1.5}
.gamecard .hs{margin-top:9px;font-size:11px;color:var(--xp);font-weight:700;min-height:14px}

.gamestage{display:flex;flex-direction:column;align-items:center;gap:12px}
.gamestage canvas{
  border-radius:var(--r);border:1px solid var(--stroke);
  max-width:100%;height:auto;background:#05070d;
}
.ghint{font-size:11.5px;color:var(--faint);text-align:center;line-height:1.7}
.ghint kbd{margin:0 2px}

/* ── chess ─────────────────────────────────────────────────────── */
.chessboard{
  display:grid;grid-template-columns:repeat(8,1fr);
  width:min(468px,86vw);aspect-ratio:1;border-radius:var(--r);
  overflow:hidden;border:1px solid var(--stroke-2);box-shadow:var(--shadow);
}
.sq{position:relative;display:grid;place-items:center;cursor:pointer;transition:background .18s}
.sq.l{background:#e4e9f2}
.sq.d{background:#6f85ad}
.sq.sel{background:#5b8cff!important}
.sq.last{box-shadow:inset 0 0 0 3px rgba(255,209,102,.6)}
.sq:hover{filter:brightness(1.09)}
.sq .pc{
  font-size:clamp(21px,4.4vw,35px);line-height:1;user-select:none;
  transition:transform .18s var(--ease);
}
.sq .pc.w{color:#fdfefe;text-shadow:0 1.5px 2px rgba(0,0,0,.55)}
.sq .pc.b{color:#181d28;text-shadow:0 1px 1px rgba(255,255,255,.28)}
.sq:hover .pc{transform:scale(1.12)}
.sq .dot{
  position:absolute;width:26%;height:26%;border-radius:50%;
  background:rgba(30,50,90,.42);pointer-events:none;
}
.sq .dot.cap{
  width:82%;height:82%;background:transparent;
  border:3.5px solid rgba(255,93,115,.75);
}
.sq .rk,.sq .fl{
  position:absolute;font-size:9px;font-style:normal;font-weight:800;
  color:rgba(20,30,50,.45);pointer-events:none;
}
.sq .rk{top:2px;left:3px}
.sq .fl{bottom:1px;right:3px}
.sq.d .rk,.sq.d .fl{color:rgba(240,245,255,.55)}

/* ── chess mode bar ─────────────────────────────────────────────── */
.chessbar{
  display:flex;gap:10px;flex-wrap:wrap;justify-content:center;
  margin-bottom:12px;width:100%;
}
.cbgroup{
  display:flex;gap:4px;padding:4px;border-radius:12px;
  background:var(--surface);border:1px solid var(--stroke);
}
.cbtn{
  padding:6px 12px;border:0;border-radius:9px;background:transparent;
  color:var(--muted);font-family:inherit;font-size:11.5px;font-weight:650;
  cursor:pointer;transition:all .22s var(--ease);white-space:nowrap;
}
.cbtn:hover{background:var(--surface-3);color:var(--text)}
.cbtn.on{background:var(--grad);color:#fff;box-shadow:0 3px 12px -4px var(--accent)}
.cbtn:active{transform:scale(.95)}

/* ── spot the liar ─────────────────────────────────────────────── */
.liar{width:100%;max-width:620px}
.case{
  padding:16px 18px;border-radius:var(--r-lg);margin-bottom:16px;
  background:linear-gradient(135deg,rgba(255,190,61,.13),rgba(255,93,115,.09));
  border:1px solid var(--stroke-2);
}
.case .tag{
  display:inline-block;font-size:10px;font-weight:800;letter-spacing:1.2px;
  text-transform:uppercase;color:var(--warn);background:rgba(255,190,61,.14);
  padding:4px 9px;border-radius:99px;margin-bottom:9px;
}
.case p{font-size:14.5px;font-weight:650;line-height:1.5}
.case small{display:block;margin-top:8px;font-size:11.5px;color:var(--muted);line-height:1.6}

.suspects{display:flex;flex-direction:column;gap:9px}
.susp{
  display:flex;align-items:center;gap:13px;padding:13px 15px;
  border-radius:var(--r);background:var(--surface);border:1px solid var(--stroke);
  cursor:pointer;transition:all .26s var(--ease);
}
.susp:hover{background:var(--surface-3);transform:translateX(5px);border-color:var(--accent)}
.susp .face{font-size:26px;flex:none}
.susp .sbody{flex:1;min-width:0}
.susp .sbody b{font-size:13.5px;font-weight:750;display:block;margin-bottom:3px}
.susp .sbody p{font-size:13px;color:var(--muted);line-height:1.5;font-style:italic}
.susp .verdict{font-size:11.5px;font-weight:800;letter-spacing:.4px;flex:none}
.susp.liar{background:rgba(255,93,115,.13);border-color:var(--bad);cursor:default;transform:none}
.susp.liar .verdict{color:var(--bad)}
.susp.honest{opacity:.55;cursor:default;transform:none}
.susp.honest .verdict{color:var(--ok)}
.liarfoot{margin-top:16px;text-align:center;font-size:12.5px;color:var(--faint)}
"""

APP_JS = """/* ══════════════════════════════════════════════════════════════════
   FaceAttend 3.2 — frontend
   ══════════════════════════════════════════════════════════════════ */
'use strict';

const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const api = async (path, opts = {}) => {
  const tok = localStorage.getItem('fa.token');
  const r = await fetch('/api' + path, {
    ...opts,
    headers: {
      'Content-Type': 'application/json',
      ...(tok ? { 'Authorization': 'Bearer ' + tok } : {}),
      ...(opts.headers || {}),
    },
  });
  if (!r.ok) throw new Error((await r.json().catch(() => ({}))).error || r.statusText);
  return r.json();
};

const HOUSE_COLOR = { Red:'#ff5d73', Blue:'#4d9dff', Green:'#2ee6a8', Orange:'#ffa23d' };

const S = {
  students: [], present: [], settings: {}, houses: {},
  page: 'home', scanning: false, sel: new Set(),
  sort: { key: 'Sl.No', dir: 1 }, q: '', fHouse: '', fStatus: '',
  notes: [], note: null, started: Date.now(),
  fx: { sound: true, confetti: true, voice: true },
  lastSeenScan: null, sounds: []
};

/* ══ TOASTS ══════════════════════════════════════════════════════ */
function toast(title, msg = '', kind = 'info', soundKey = null) {
  const ic = { info:['💡','i-blue'], ok:['✓','i-green'], warn:['⚠','i-amber'],
               bad:['✕','i-red'], xp:['⚡','i-purple'] }[kind] || ['💡','i-blue'];
  const el = document.createElement('div');
  el.className = 'toast';
  el.style.position = 'relative';
  el.style.overflow = 'hidden';
  el.innerHTML = `<div class="ti ${ic[1]}">${ic[0]}</div>
    <div style="flex:1"><b></b>${msg ? '<p></p>' : ''}</div>
    <div class="bar"></div>`;
  el.querySelector('b').textContent = title;
  if (msg) el.querySelector('p').textContent = msg;
  $('#toasts').appendChild(el);
  setTimeout(() => { el.classList.add('out'); setTimeout(() => el.remove(), 350); }, 4000);
  // per-kind cue, mirroring the desktop build's sound_success / sound_error /
  // sound_negative — every toast (add/edit/delete/export/error/warning) now
  // has a matching chime, not just scans, badges and games. Pass soundKey
  // to use a specific real file (e.g. 'negative' for a successful delete,
  // 'export' for exports/imports/backups) instead of the kind default.
  const key = soundKey || ({ ok: 'success', bad: 'error' })[kind] || null;
  if (key) playSound(key, kind === 'bad' ? [300, 220] : [660, 880], kind === 'bad' ? .12 : .08);
  else if (kind === 'warn' && typeof beep === 'function') beep([440, 370], .09);
}

/* ══ REAL SOUND FILES (falls back to beep() if not present on disk) ═══ */
const _audioCache = {};
function playSound(key, fallbackFreqs, fallbackDur = 0.12) {
  if (!S.fx.sound) return;
  if ((S.sounds || []).includes(key)) {
    try {
      let a = _audioCache[key];
      if (!a) { a = _audioCache[key] = new Audio('/api/sound/' + key); }
      a.currentTime = 0;
      a.play().catch(() => { if (fallbackFreqs) beep(fallbackFreqs, fallbackDur); });
      return;
    } catch (e) {}
  }
  if (fallbackFreqs) beep(fallbackFreqs, fallbackDur);
}

/* ══ SOUND (WebAudio — no files needed) ══════════════════════════ */
let AC = null;
function beep(freqs = [660, 880], dur = 0.12) {
  if (!S.fx.sound) return;
  try {
    AC = AC || new (window.AudioContext || window.webkitAudioContext)();
    freqs.forEach((f, i) => {
      const o = AC.createOscillator(), g = AC.createGain();
      o.type = 'sine'; o.frequency.value = f;
      o.connect(g); g.connect(AC.destination);
      const t = AC.currentTime + i * dur;
      g.gain.setValueAtTime(0, t);
      g.gain.linearRampToValueAtTime(0.14, t + 0.015);
      g.gain.exponentialRampToValueAtTime(0.001, t + dur);
      o.start(t); o.stop(t + dur);
    });
  } catch (e) {}
}

/* ══ CONFETTI ════════════════════════════════════════════════════ */
function confetti(n = 55) {
  if (!S.fx.confetti) return;
  const box = $('#conf');
  const cols = ['#5b8cff','#a56bff','#22d3ee','#2ee6a8','#ffd166','#ff5d73'];
  for (let i = 0; i < n; i++) {
    const p = document.createElement('div');
    p.className = 'cf';
    p.style.left = Math.random() * 100 + 'vw';
    p.style.top = '-14px';
    p.style.background = cols[(Math.random() * cols.length) | 0];
    p.style.borderRadius = Math.random() > .5 ? '50%' : '2px';
    p.style.animationDuration = (1.9 + Math.random() * 1.5) + 's';
    p.style.animationDelay = (Math.random() * .35) + 's';
    p.style.opacity = .55 + Math.random() * .45;
    box.appendChild(p);
    setTimeout(() => p.remove(), 3800);
  }
}

/* ══ NAV ═════════════════════════════════════════════════════════ */
function go(page) {
  const shell = $('#shell'), panel = $('#panel'), main = $('#main');
  const section = $(`.page[data-page="${page}"]`);
  if (!section) return;

  // FaceAttend 3.2 gives Home its two-column dashboard and lets every
  // other tab breathe across the full content area. Moving the active
  // section keeps one source of truth for every control/id.
  if (page === 'home') {
    if (section.parentElement !== panel) panel.appendChild(section);
    shell.classList.remove('wide');
    $('#viewDash').style.display = '';
  } else {
    if (section.parentElement !== main) main.appendChild(section);
    shell.classList.add('wide');
    $('#viewDash').style.display = 'none';
  }

  S.page = page;
  $$('.navbtn').forEach(b => b.classList.toggle('on', b.dataset.page === page));
  $$('.page').forEach(p => p.classList.toggle('on', p === section));
  section.classList.toggle('wide-page', page !== 'home');
  if (page === 'analytics') drawAnalytics();
  if (page === 'analytics') drawHouses();
  if (page === 'notes')     loadNotes();
  if (page === 'mail')      fillMailStudents();
  if (page === 'awards')    refreshAwards();
  if (page === 'settings')  refreshDemo();
  if (page === 'console') $('#cmdInput')?.focus();
  localStorage.setItem('fa.page', page);
}

function showAnalyticsPane(which) {
  const analytics = $('#analyticsPane'), houses = $('#housesPane');
  if (!analytics || !houses) return;
  const isHouses = which === 'houses';
  analytics.style.display = isHouses ? 'none' : '';
  houses.style.display = isHouses ? '' : 'none';
  $('#analyticsTab')?.classList.toggle('on', !isHouses);
  $('#housesTab')?.classList.toggle('on', isHouses);
  if (isHouses) drawHouses(); else drawAnalytics();
}

function showNotepadPane(which) {
  const notes = $('#notesPane'), board = $('#boardPane');
  if (!notes || !board) return;
  const isBoard = which === 'whiteboard';
  notes.style.display = isBoard ? 'none' : '';
  board.style.display = isBoard ? '' : 'none';
  $('#notesTab')?.classList.toggle('on', !isBoard);
  $('#boardTab')?.classList.toggle('on', isBoard);
  if (isBoard) initBoard();
}

/* ══ RENDER: KPIs ════════════════════════════════════════════════ */
function countUp(el, to, dur = 900) {
  const from = parseFloat(el.dataset.v || 0);
  if (from === to) { el.textContent = fmt(to); return; }
  const t0 = performance.now();
  const tick = t => {
    const k = Math.min(1, (t - t0) / dur);
    const e = 1 - Math.pow(1 - k, 3);
    el.textContent = fmt(from + (to - from) * e);
    if (k < 1) requestAnimationFrame(tick); else el.dataset.v = to;
  };
  requestAnimationFrame(tick);
}
const fmt = n => (Math.round(n * 10) / 10 % 1 === 0) ? String(Math.round(n)) : n.toFixed(1);

function drawKpis() {
  const total = S.students.length;
  const present = S.present.length;
  const absent = Math.max(0, total - present);
  const top = [...S.students].sort((a, b) => (+b.XP || 0) - (+a.XP || 0))[0];
  const totalXp = S.students.reduce((s, x) => s + (+x.XP || 0), 0);
  const rate = total ? (present / total * 100) : 0;

  const cards = [
    { ic:'🎓', cls:'i-blue',   lb:'Total students', v:total,   ft:'Enrolled' },
    { ic:'✓',  cls:'i-green',  lb:'Present today',  v:present, ft:`${rate.toFixed(0)}% attendance` },
    { ic:'⏰', cls:'i-amber',  lb:'Absent today',   v:absent,  ft:'Not yet seen' },
    { ic:'★',  cls:'i-purple', lb:'XP leader',      v:(top ? +top.XP || 0 : 0),
      ft:(top ? top.Name : '—'), suffix:' XP' },
    { ic:'⚡', cls:'i-red',    lb:'Total XP',       v:totalXp, ft:'Across roster' }
  ];

  const box = $('#kpis');
  if (box.children.length !== cards.length) {
    box.innerHTML = cards.map(c => `
      <div class="kpi">
        <div class="ic ${c.cls}">${c.ic}</div>
        <div class="lb">${c.lb}</div>
        <div class="vl"><span class="n">0</span><span class="sfx"></span></div>
        <div class="ft"></div>
      </div>`).join('');
  }
  [...box.children].forEach((el, i) => {
    countUp(el.querySelector('.n'), cards[i].v);
    el.querySelector('.sfx').textContent = cards[i].suffix || '';
    el.querySelector('.ft').textContent = cards[i].ft;
  });
}

/* ══ RENDER: TABLE ═══════════════════════════════════════════════ */
const COLS = [
  { k:'Sl.No', t:'#', w:'46px' },
  { k:'Name',  t:'Student' },
  { k:'Roll No', t:'Roll' },
  { k:'UID',   t:'UID' },
  { k:'Grade', t:'Grade' },
  { k:'House', t:'House' },
  { k:'_status', t:'Today' },
  { k:'No. of Days Present', t:'Days' },
  { k:'XP', t:'XP' },
  { k:'Streak', t:'Streak' },
  { k:'Parent Phone', t:'Phone' },
  { k:'Parent Email', t:'Email' }
];

function visible() {
  const q = S.q.toLowerCase().trim();
  return S.students.filter(s => {
    if (S.fHouse && s.House !== S.fHouse) return false;
    const isP = S.present.includes(String(s.UID));
    if (S.fStatus === 'present' && !isP) return false;
    if (S.fStatus === 'absent'  &&  isP) return false;
    if (!q) return true;
    return COLS.some(c => String(s[c.k] ?? '').toLowerCase().includes(q));
  }).sort((a, b) => {
    const k = S.sort.key;
    let x = a[k], y = b[k];
    if (k === '_status') { x = S.present.includes(String(a.UID)); y = S.present.includes(String(b.UID)); }
    const nx = parseFloat(x), ny = parseFloat(y);
    if (!isNaN(nx) && !isNaN(ny)) return (nx - ny) * S.sort.dir;
    return String(x ?? '').localeCompare(String(y ?? '')) * S.sort.dir;
  });
}

function drawTable(flashUid = null) {
  $('#thead').innerHTML = COLS.map(c =>
    `<th data-k="${c.k}" ${c.w ? `style="width:${c.w}"` : ''}>${c.t}
      <span class="ar">${S.sort.key === c.k ? (S.sort.dir > 0 ? '▲' : '▼') : '⇅'}</span></th>`).join('');

  const rows = visible();
  const tb = $('#tbody');

  if (!rows.length) {
    tb.innerHTML = `<tr><td colspan="${COLS.length}">
      <div class="empty"><div class="em">${S.students.length ? '🔍' : '🎓'}</div>
      <h3>${S.students.length ? 'No students match that search' : 'No students enrolled yet'}</h3>
      <p style="font-size:13px">${S.students.length ? 'Try a different name, roll number or house.' : 'Press Add to enrol your first student.'}</p>
      </div></td></tr>`;
    return;
  }

  tb.innerHTML = rows.map(s => {
    const uid = String(s.UID ?? '');
    const isP = S.present.includes(uid);
    const initials = String(s.Name || '?').split(/\\s+/).map(w => w[0]).slice(0, 2).join('').toUpperCase();
    const hc = HOUSE_COLOR[s.House] || 'var(--faint)';
    const streak = +s.Streak || 0;
    return `<tr data-uid="${uid}" class="${S.sel.has(uid) ? 'sel' : ''} ${flashUid === uid ? 'flash' : ''}">
      <td style="color:var(--faint);font-size:12px">${s['Sl.No'] ?? ''}</td>
      <td><div class="who">
        <img class="av" src="/api/photo/${encodeURIComponent(uid)}" alt=""
             onerror="this.outerHTML='<div class=&quot;av ph&quot; style=&quot;background:${hc}&quot;>${initials}</div>'">
        <div><b>${esc(s.Name)}</b><span>${esc(s.Grade || '')}</span></div>
      </div></td>
      <td>${esc(s['Roll No'])}</td>
      <td style="font-family:var(--mono);font-size:12px;color:var(--muted)">${esc(uid)}</td>
      <td>${esc(s.Grade)}</td>
      <td><span class="hd" style="background:${hc}"></span>${esc(s.House)}</td>
      <td><span class="pill ${isP ? 'on' : 'off'}">${isP ? '✓ Present' : '· Absent'}</span></td>
      <td style="font-variant-numeric:tabular-nums">${s['No. of Days Present'] ?? 0}</td>
      <td><span class="pill xp">⚡ ${s.XP ?? 0}</span></td>
      <td>${streak > 0 ? `🔥 ${streak}` : '<span style="color:var(--faint)">—</span>'}</td>
      <td style="color:var(--muted);font-size:12.5px">${esc(s['Parent Phone'])}</td>
      <td style="color:var(--muted);font-size:12.5px">${esc(s['Parent Email'])}</td>
    </tr>`;
  }).join('');
}

const esc = v => String(v ?? '').replace(/[&<>"']/g, c =>
  ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));

/* ══ DATA ════════════════════════════════════════════════════════ */
async function refresh(flashUid = null) {
  try {
    const d = await api('/state');
    S.students = d.students; S.present = d.present;
    S.settings = d.settings; S.houses = d.houses;
    S.scanning = d.scanning;
    S.trend = d.trend || [];
    S.sounds = d.sounds || S.sounds || [];
    drawKpis(); drawTable(flashUid);
    $('#scanDot').classList.toggle('live', d.scanning);
    paintWindow(d.attendance);
    applyRole(d);

    // Tell the user plainly when the scanner cannot actually recognise
    // anyone. Before this it looked like it was working and simply never
    // marked a soul.
    const warn = $('#recogWarn'), msg = $('#recogWarnMsg');
    if (warn && msg) {
      if (!d.face_engine) {
        msg.innerHTML = 'The face recognition engine is not installed, so the '
          + 'camera shows video but cannot identify anyone.<br><br>'
          + 'Install it once (takes 10–25 minutes, it compiles dlib):<br>'
          + '<span style="font-family:var(--mono);color:var(--text)">'
          + 'pip install cmake dlib face_recognition</span>';
        warn.style.display = '';
      } else if (!d.enrolled_faces) {
        msg.textContent = 'No student has a face photo yet. Use Add (or Edit) '
          + 'and capture 3 photos so the scanner has something to match against.';
        warn.style.display = '';
      } else {
        warn.style.display = 'none';
      }
    }
    $('#stScans').textContent = d.present.length;
    $('#stUnknown').textContent = d.unknown_count ?? 0;
    if (S.page === 'analytics') drawAnalytics();
    if (S.page === 'houses') drawHouses();
    if (!S._missingWarned && (d.missing || []).includes('face_recognition')) {
      S._missingWarned = true;
      toast('Face recognition engine missing',
        'Scans will never match a student. Restart the app once so it can finish installing "face_recognition", or run: pip install face_recognition',
        'bad');
    }
  } catch (e) { toast('Could not load data', e.message, 'bad'); }
}

/* ══ SCANNER ═════════════════════════════════════════════════════ */
async function toggleScan() {
  const btn = $('#btnScan');
  btn.disabled = true;
  try {
    if (S.scanning) {
      await api('/scan/stop', { method:'POST' });
      S.scanning = false;
      $('#camFeed').style.display = 'none';
      $('#camFeed').src = '';
      $('#camOff').style.display = '';
      $('#camBadge').style.display = 'none';
      btn.textContent = '▶ Start Scan';
      btn.className = 'btn primary';
      toast('Scanner stopped', '', 'info');
    } else {
      const r = await api('/scan/start', { method:'POST' });
      if (!r.ok) throw new Error(r.error || 'Camera unavailable');
      S.scanning = true;
      $('#camOff').style.display = 'none';
      $('#camFeed').style.display = '';
      $('#camFeed').src = '/api/stream?t=' + Date.now();
      $('#camBadge').style.display = '';
      btn.textContent = '■ Stop Scan';
      btn.className = 'btn bad';
      toast('Scanner started', 'Watching for faces…', 'info');
      playSound('camera_start', [520, 700]);
    }
    $('#scanDot').classList.toggle('live', S.scanning);
  } catch (e) {
    toast('Camera error', e.message, 'bad');
  }
  btn.disabled = false;
}

/* poll for recognitions while scanning */
setInterval(async () => {
  if (!S.scanning) return;
  try {
    const d = await api('/scan/last');
    if (d.stamp && d.stamp !== S.lastSeenScan) {
      S.lastSeenScan = d.stamp;
      if (d.name) {
        $('#recog').classList.remove('hit');
        void $('#recog').offsetWidth;
        $('#recog').classList.add('hit');
        $('#recogName').textContent = `Welcome, ${d.name}!`;
        $('#recogMeta').innerHTML =
          `<span class="pill ${d.late ? 'late' : 'on'}">${d.late ? '⏰ Late' : '✓ On time'}</span>
           <span class="pill xp">⚡ +${d.xp} XP</span>
           <span style="color:var(--faint)">${d.time}</span>`;
        addFeed(d.time, `${d.name} marked ${d.late ? 'late' : 'present'} · +${d.xp} XP`);
        toast(`${d.name} recognised`,
              `+${d.xp} XP · ${d.label || (d.late ? 'late' : 'on time')}`, 'xp');
        speak(`Welcome, ${d.name}`);
        (d.badges || []).forEach(b => {
          toast('🏅 Badge unlocked', b, 'xp');
          confetti(70);
        });
        beep([740, 990, 1240]);
        confetti();
        refresh(String(d.uid));
      }
    }
  } catch (e) {}
}, 1400);

function addFeed(t, msg) {
  const el = document.createElement('div');
  el.className = 'fi';
  el.innerHTML = `<time>${t}</time><span></span>`;
  el.querySelector('span').textContent = msg;
  $('#feed').prepend(el);
  while ($('#feed').children.length > 60) $('#feed').lastChild.remove();
  logLine(t, msg, 's');
}

function logLine(t, msg, kind = 'i') {
  const el = document.createElement('div');
  el.className = 'ln';
  el.innerHTML = `<span class="t">${t}</span><span class="${kind}"></span>`;
  el.querySelector('span:last-child').textContent = msg;
  $('#consoleLog').appendChild(el);
  while ($('#consoleLog').children.length > 400) $('#consoleLog').lastChild.remove();
  $('#consoleLog').scrollTop = $('#consoleLog').scrollHeight;
}

async function runConsoleCommand(value = null) {
  const input = $('#cmdInput');
  const raw = (value ?? input?.value ?? '').trim();
  if (!raw) return;
  if (input) input.value = '';
  logLine(now(), raw, 'cmd');
  try {
    const d = await api('/command', { method:'POST', body:JSON.stringify({ command:raw }) });
    if (d.clear) $('#consoleLog').innerHTML = '';
    (d.lines || []).forEach(line => logLine(line.time || now(), line.text, line.kind || 'i'));
    if (d.refresh) refresh();
    if (d.navigate === 'arcade') openArcade();
    else if (d.navigate && d.navigate.startsWith('game:')) openArcade(d.navigate.slice(5));
    else if (d.navigate === 'houses') { go('analytics'); showAnalyticsPane('houses'); }
    else if (d.navigate === 'board') { go('notes'); showNotepadPane('whiteboard'); }
    else if (d.navigate) go(d.navigate);
    // 3.2 console actions
    if (d.open === 'certificates') { go('awards'); setTimeout(makeCertificates, 260); }
    if (d.open === 'calendar') openCalendar();
    if (d.qr) showQr(d.qr);
    if (d.notify_absent) notifyAbsent();
    if (d.toggle_theme) { const t = $('#swTheme'); if (t) t.click(); }
    if (d.reload) location.reload();
    if (d.unlock_dev) unlockDev();
  } catch (e) {
    logLine(now(), 'command failed: ' + e.message, 'e');
  }
}

/* ══ ANALYTICS ═══════════════════════════════════════════════════ */
function drawAnalytics() {
  const total = S.students.length;
  const rate = total ? (S.present.length / total * 100) : 0;
  $('#avgAtt').textContent = rate.toFixed(1) + '%';
  const maxDays = Math.max(0, ...S.students.map(s => +s['No. of Days Present'] || 0));
  $('#schoolDays').textContent = maxDays;

  // house bars
  $('#houseBars').innerHTML = Object.entries(HOUSE_COLOR).map(([h, c]) => {
    const st = S.houses[h] || { pct:0 };
    return `<div class="bcol">
      <div class="bk" data-v="${st.pct.toFixed(0)}%" style="height:0;background:linear-gradient(180deg,${c},${c}66)"></div>
      <div class="bl" style="color:${c}">${h}</div>
    </div>`;
  }).join('');
  const names = Object.keys(HOUSE_COLOR);
  requestAnimationFrame(() => {
    $$('#houseBars .bk').forEach((b, i) => {
      const h = S.houses[names[i]]?.pct || 0;
      b.style.height = Math.max(3, h) + '%';
    });
  });

  // sparkline trend — real attendance history from the audit log
  drawTrend(S.trend || []);
  drawTopXp();
}

function drawTrend(pts) {
  if (!pts.length) {
    $('#trend').innerHTML =
      '<text x="150" y="70" text-anchor="middle" fill="#5b6880" font-size="11">No history yet</text>';
    return;
  }
  const n = pts.length;
  const d = pts.map((v, i) => `${n === 1 ? 150 : (i / (n - 1)) * 300},${130 - (v / 100) * 118}`).join(' ');
  $('#trend').innerHTML = `
    <defs><linearGradient id="tg" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#5b8cff" stop-opacity=".45"/>
      <stop offset="100%" stop-color="#5b8cff" stop-opacity="0"/>
    </linearGradient></defs>
    <polygon points="0,130 ${d} 300,130" fill="url(#tg)"/>
    <polyline points="${d}" fill="none" stroke="#5b8cff" stroke-width="2.2"
      stroke-linejoin="round" stroke-linecap="round"/>`;
}

function drawTopXp() {
  // top XP
  const top = [...S.students].sort((a, b) => (+b.XP || 0) - (+a.XP || 0)).slice(0, 8);
  $('#topXp').innerHTML = top.length ? top.map((s, i) => `
    <div class="lbrow p${i + 1}">
      <div class="rk">${i + 1}</div>
      <div class="nm">${esc(s.Name)}<small>${esc(s.Grade || '')} · ${esc(s.House || '')}</small></div>
      <span class="pill xp">⚡ ${s.XP ?? 0}</span>
    </div>`).join('')
    : '<p style="color:var(--faint);font-size:13px">No students yet.</p>';
}

function drawHouses() {
  const entries = Object.entries(HOUSE_COLOR);
  const ranked = [...entries].sort((a, b) => (S.houses[b[0]]?.pct || 0) - (S.houses[a[0]]?.pct || 0));
  const leader = ranked[0]?.[0];

  $('#houseCards').innerHTML = entries.map(([h, c]) => {
    const st = S.houses[h] || { pct:0, present:0, students:0 };
    return `<div class="hcard ${h === leader && st.pct > 0 ? 'lead' : ''}"
      style="background:linear-gradient(160deg,${c}18,transparent);border-color:${c}44">
      <div class="hn"><span class="hd" style="background:${c}"></span>${h}</div>
      <div class="hp" style="color:${c}">${st.pct.toFixed(1)}%</div>
      <div class="hs">${st.present} / ${st.students} present</div>
      <div class="hbar" style="background:${c};width:${st.pct}%"></div>
    </div>`;
  }).join('');

  $('#houseRank').innerHTML = ranked.map(([h, c], i) => {
    const st = S.houses[h] || { pct:0, present:0, students:0 };
    return `<div class="lbrow p${i + 1}">
      <div class="rk">${i + 1}</div>
      <div class="nm">${h}<small>${st.present}/${st.students} students</small></div>
      <b style="color:${c};font-variant-numeric:tabular-nums">${st.pct.toFixed(1)}%</b>
    </div>`;
  }).join('');
}

/* ══ MODALS ══════════════════════════════════════════════════════ */
function modal(html) {
  $('#modal').innerHTML = html;
  $('#veil').classList.add('on');
}
const closeModal = () => $('#veil').classList.remove('on');

function studentForm(s = null) {
  const f = (k, v = '') => esc(s ? (s[k] ?? '') : v);
  modal(`
    <h2>${s ? 'Edit student' : 'Add a student'}</h2>
    <p class="sub">${s ? 'Update the details below.' : 'Fill in the details, then capture a face photo.'}</p>
    <div class="grid2">
      <div class="field"><label>Full name *</label><input id="fName" value="${f('Name')}"></div>
      <div class="field"><label>Roll number</label><input id="fRoll" value="${f('Roll No')}"></div>
    </div>
    <div class="grid2">
      <div class="field"><label>Grade / section</label><input id="fGrade" value="${f('Grade')}"></div>
      <div class="field"><label>House</label><select id="fHouseSel">
        ${['Red','Blue','Green','Orange'].map(h =>
          `<option ${s && s.House === h ? 'selected' : ''}>${h}</option>`).join('')}
      </select></div>
    </div>
    <div class="grid2">
      <div class="field"><label>Parent phone</label><input id="fPhone" value="${f('Parent Phone')}"></div>
      <div class="field"><label>Parent email</label><input id="fEmail" value="${f('Parent Email')}"></div>
    </div>
    ${s ? '' : `<div class="field">
      <label>Face photos — capture 3 or more from different angles</label>
      <div class="row">
        <button class="btn" id="fCapture">📷 Capture from camera</button>
        <label class="btn" style="cursor:pointer">📁 Upload
          <input type="file" id="fFile" accept="image/*" style="display:none"></label>
      </div>
      <div id="fPreview" style="margin-top:11px"></div>
    </div>`}
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="fSave">${s ? 'Save changes' : 'Add student'}</button>
    </div>`);

  let photo = null;
  let extraShots = [];
  if (!s) {
    $('#fCapture').onclick = async () => {
      try {
        toast('Capturing…', 'Look at the camera', 'info');
        const r = await api('/capture', { method:'POST' });
        if (!r.ok) throw new Error(r.error);
        if (photo === null) photo = r.image;
        else extraShots.push(r.image);
        const all = [photo, ...extraShots];
        $('#fPreview').innerHTML = all.map(src =>
          `<img src="${src}" style="width:74px;height:74px;object-fit:cover;
             border-radius:11px;border:2px solid var(--accent);margin-right:6px">`).join('')
          + `<div style="font-size:11.5px;color:${all.length >= 3 ? 'var(--ok)' : 'var(--muted)'};
                margin-top:7px">${all.length} photo(s) captured${
                all.length < 3 ? ' — 3 or more gives much better recognition'
                               : ' — good coverage'}</div>`;
        toast(`Photo ${all.length} captured`,
              all.length < 3 ? 'Turn your head slightly and capture again' : '', 'ok');
      } catch (e) { toast('Capture failed', e.message, 'bad'); }
    };
    $('#fFile').onchange = e => {
      const file = e.target.files[0]; if (!file) return;
      const rd = new FileReader();
      rd.onload = () => {
        photo = rd.result;
        $('#fPreview').innerHTML =
          `<img src="${photo}" style="width:100px;height:100px;object-fit:cover;border-radius:13px;border:2px solid var(--accent)">`;
      };
      rd.readAsDataURL(file);
    };
  }

  $('#fSave').onclick = async () => {
    const body = {
      Name: $('#fName').value.trim(),
      'Roll No': $('#fRoll').value.trim(),
      Grade: $('#fGrade').value.trim(),
      House: $('#fHouseSel').value,
      'Parent Phone': $('#fPhone').value.trim(),
      'Parent Email': $('#fEmail').value.trim()
    };
    if (!body.Name) return toast('Name is required', '', 'warn');
    try {
      if (s) {
        await api('/students/' + encodeURIComponent(s.UID), { method:'PUT', body:JSON.stringify(body) });
        toast('Student updated', body.Name, 'ok');
      } else {
        body.photo = photo;
        body.extra_photos = extraShots;
        const r = await api('/students', { method:'POST', body:JSON.stringify(body) });
        toast('Student added', `${body.Name} · UID ${r.uid}`, 'ok');
        confetti(30); beep([620, 820, 1040]);
      }
      closeModal(); refresh();
    } catch (e) { toast('Could not save', e.message, 'bad'); }
  };
}

/* ══ WHITEBOARD ══════════════════════════════════════════════════ */
let boardInit = false, drawing = false, brushCol = '#111827', brushSize = 4, erasing = false;
function initBoard() {
  const cv = $('#board'), cx = cv.getContext('2d');
  if (!boardInit) {
    boardInit = true;
    cx.fillStyle = '#fff'; cx.fillRect(0, 0, cv.width, cv.height);
    cx.lineCap = cx.lineJoin = 'round';

    const COLORS = ['#111827','#ff5d73','#4d9dff','#2ee6a8','#ffa23d','#a56bff','#22d3ee','#ffffff'];
    $('#swatches').innerHTML = COLORS.map((c, i) =>
      `<div class="sw2 ${i === 0 ? 'on' : ''}" data-c="${c}" style="background:${c};${c === '#ffffff' ? 'border-color:var(--stroke-2)' : ''}"></div>`).join('');
    $$('.sw2').forEach(sw => sw.onclick = () => {
      $$('.sw2').forEach(x => x.classList.remove('on'));
      sw.classList.add('on'); brushCol = sw.dataset.c; erasing = false;
    });

    const pos = e => {
      const r = cv.getBoundingClientRect();
      const p = e.touches ? e.touches[0] : e;
      return { x:(p.clientX - r.left) * (cv.width / r.width), y:(p.clientY - r.top) * (cv.height / r.height) };
    };
    const start = e => { drawing = true; const p = pos(e); cx.beginPath(); cx.moveTo(p.x, p.y); e.preventDefault(); };
    const move = e => {
      if (!drawing) return;
      const p = pos(e);
      cx.strokeStyle = erasing ? '#ffffff' : brushCol;
      cx.lineWidth = erasing ? brushSize * 3 : brushSize;
      cx.lineTo(p.x, p.y); cx.stroke(); e.preventDefault();
    };
    const end = () => drawing = false;
    cv.addEventListener('mousedown', start); cv.addEventListener('mousemove', move);
    window.addEventListener('mouseup', end);
    cv.addEventListener('touchstart', start); cv.addEventListener('touchmove', move);
    cv.addEventListener('touchend', end);

    $('#brush').oninput = e => { brushSize = +e.target.value; $('#brushVal').textContent = brushSize; };
    $('#btnEraser').onclick = () => {
      erasing = !erasing;
      $('#btnEraser').classList.toggle('warn', erasing);
      toast(erasing ? 'Eraser on' : 'Eraser off', '', 'info');
    };
    $('#btnClear').onclick = () => {
      cx.fillStyle = '#fff'; cx.fillRect(0, 0, cv.width, cv.height);
      toast('Board cleared', '', 'info');
    };
    $('#btnSaveBoard').onclick = async () => {
      try {
        await api('/board', { method:'POST', body:JSON.stringify({ image: cv.toDataURL('image/png') }) });
        toast('Board saved', 'Whiteboard folder', 'ok');
      } catch (e) { toast('Save failed', e.message, 'bad'); }
    };
  }
}

/* ══ NOTES ═══════════════════════════════════════════════════════ */
async function loadNotes() {
  try {
    S.notes = (await api('/notes')).notes;
    $('#noteList').innerHTML = S.notes.length ? S.notes.map(n => `
      <div class="noteitem ${S.note === n.name ? 'on' : ''}" data-n="${esc(n.name)}">
        <span>${esc(n.name.replace(/\\.txt$/, ''))}</span><small>${n.size}b</small>
      </div>`).join('')
      : '<p style="color:var(--faint);font-size:12.5px;padding:8px">No notes yet.</p>';
    $$('#noteList .noteitem').forEach(el => el.onclick = () => openNote(el.dataset.n));
  } catch (e) {}
}
async function openNote(name) {
  const d = await api('/notes/' + encodeURIComponent(name));
  S.note = name;
  $('#noteTitle').textContent = name.replace(/\\.txt$/, '');
  $('#noteBody').value = d.body;
  loadNotes();
}
async function saveNote() {
  const title = ($('#noteTitle').textContent || 'Untitled note').trim();
  try {
    await api('/notes', { method:'POST', body:JSON.stringify({ name:title, body:$('#noteBody').value }) });
    S.note = title + '.txt';
    if (window.__noteClean) window.__noteClean();
    const t = $('#noteTitle');
    if (t) t.textContent = t.textContent.replace(/ \u2022$/, '');
    toast('Note saved', title, 'ok');
    loadNotes();
  } catch (e) { toast('Save failed', e.message, 'bad'); }
}

/* ══ ACTIONS ═════════════════════════════════════════════════════ */
async function doExport() {
  try {
    const r = await api('/export', { method:'POST' });
    toast('Exported', r.path, 'ok', 'export');
    logLine(now(), 'Exported roster to ' + r.path, 's');
  } catch (e) { toast('Export failed', e.message, 'bad'); }
}
async function doBackup() {
  try {
    const r = await api('/backup', { method:'POST' });
    toast('Backup saved', r.path, 'ok', 'export');
  } catch (e) { toast('Backup failed', e.message, 'bad'); }
}
const now = () => new Date().toLocaleTimeString('en-GB');

function selectedStudents() {
  return S.students.filter(s => S.sel.has(String(s.UID)));
}

/* ══ AI BUDDY ════════════════════════════════════════════════════ */
function bubble(text, who) {
  const t = new Date().toLocaleTimeString([], { hour:'2-digit', minute:'2-digit' });
  const turn = document.createElement('div');
  turn.className = 'turn ' + who;
  turn.innerHTML = `<div class="pic">${who === 'bot' ? '\\u{1F916}' : 'You'}</div>
    <div class="body"><div class="msg"></div><div class="when">${t}</div></div>`;
  turn.querySelector('.msg').textContent = text;
  $('#chat').appendChild(turn);
  $('#chat').scrollTop = $('#chat').scrollHeight;
  return turn;
}
async function ask(q) {
  if (!q.trim()) return;
  bubble(q, 'you');
  $('#botInput').value = '';
  const think = document.createElement('div');
  think.className = 'turn bot';
  think.innerHTML = '<div class="pic">\\u{1F916}</div><div class="body">' +
                    '<div class="msg think"><i></i><i></i><i></i></div></div>';
  $('#chat').appendChild(think);
  $('#chat').scrollTop = $('#chat').scrollHeight;
  $('#botSend').disabled = true;
  try {
    const d = await api('/ask', { method:'POST', body:JSON.stringify({ q }) });
    think.remove();
    bubble(d.answer, 'bot');
    setBotMode(d.mode);
    if (d.refresh) refresh();
    if (d.navigate === 'arcade') openArcade();
    else if (d.navigate && d.navigate.startsWith('game:')) openArcade(d.navigate.slice(5));
    else if (d.navigate) go(d.navigate);
  } catch (e) {
    think.remove();
    bubble('Something went wrong: ' + e.message, 'bot');
  }
  $('#botSend').disabled = false;
  $('#botInput').focus();
}

function setBotMode(mode) {
  const el = $('#botStatus');
  if (!el) return;
  const ai = mode === 'AI mode';
  el.classList.toggle('ai', ai);
  el.querySelector('span:last-child').innerHTML = ai
    ? 'Reading live roster · <b>AI mode</b>'
    : 'Reading live roster · <b>offline brain</b>';
}

/* ══ MAIL ════════════════════════════════════════════════════════ */
function fillMailStudents() {
  $('#mailStudent').innerHTML = '<option value="">Select a student…</option>' +
    S.students.map(s => `<option value="${esc(s.UID)}">${esc(s.Name)} — ${esc(s.Grade || '')}</option>`).join('');
}

/* ══ INIT ════════════════════════════════════════════════════════ */
function init() {
  // clock
  const tick = () => {
    const d = new Date();
    $('#clock').textContent = d.toLocaleTimeString('en-GB');
    $('#today').textContent = d.toLocaleDateString(undefined,
      { weekday:'long', day:'numeric', month:'long', year:'numeric' }) + ' ·';
    const mins = Math.floor((Date.now() - S.started) / 60000);
    $('#stUptime').textContent = mins < 60 ? mins + 'm' : Math.floor(mins / 60) + 'h ' + (mins % 60) + 'm';
  };
  tick(); setInterval(tick, 1000);

  // nav
  $$('.navbtn').forEach(b => b.onclick = () => go(b.dataset.page));
  const remembered = localStorage.getItem('fa.page') || 'home';
  go(['home','analytics','bot','mail','console','notes','settings'].includes(remembered)
    ? remembered : 'home');

  // ripples
  document.addEventListener('click', e => {
    const b = e.target.closest('.btn');
    if (!b) return;
    const r = document.createElement('span');
    r.className = 'ripple';
    const rect = b.getBoundingClientRect();
    const size = Math.max(rect.width, rect.height);
    r.style.width = r.style.height = size + 'px';
    r.style.left = (e.clientX - rect.left - size / 2) + 'px';
    r.style.top  = (e.clientY - rect.top - size / 2) + 'px';
    b.appendChild(r); setTimeout(() => r.remove(), 620);
  });

  // scanner
  $('#btnScan').onclick = toggleScan;
  $('#btnRefresh').onclick = () => { refresh(); toast('Refreshed', '', 'info'); };

  // table interactions
  $('#q').oninput = e => { S.q = e.target.value; drawTable(); };
  $('#fHouse').onchange = e => { S.fHouse = e.target.value; drawTable(); };
  $('#fStatus').onchange = e => { S.fStatus = e.target.value; drawTable(); };
  $('#thead').onclick = e => {
    const th = e.target.closest('th'); if (!th) return;
    const k = th.dataset.k;
    S.sort = { key:k, dir: S.sort.key === k ? -S.sort.dir : 1 };
    drawTable();
  };
  $('#tbody').onclick = e => {
    const tr = e.target.closest('tr'); if (!tr || !tr.dataset.uid) return;
    const uid = tr.dataset.uid;
    if (e.shiftKey || S.sel.size) {
      S.sel.has(uid) ? S.sel.delete(uid) : S.sel.add(uid);
    } else {
      S.sel.clear(); S.sel.add(uid);
    }
    drawTable();
  };
  $('#tbody').ondblclick = e => {
    const tr = e.target.closest('tr'); if (!tr) return;
    const s = S.students.find(x => String(x.UID) === tr.dataset.uid);
    if (s) studentForm(s);
  };
  $('#btnSelAll').onclick = () => { visible().forEach(s => S.sel.add(String(s.UID))); drawTable(); };
  $('#btnClearSel').onclick = () => { S.sel.clear(); drawTable(); };

  // CRUD
  $('#btnAdd').onclick = () => studentForm();
  $('#btnEdit').onclick = () => {
    const sel = selectedStudents();
    if (sel.length !== 1) return toast('Select exactly one student', 'Click a row first', 'warn');
    studentForm(sel[0]);
  };
  $('#btnRemove').onclick = () => {
    const sel = selectedStudents();
    if (!sel.length) return toast('Nothing selected', 'Click one or more rows', 'warn');
    modal(`<h2>Remove ${sel.length} student${sel.length > 1 ? 's' : ''}?</h2>
      <p class="sub">This deletes their record and face data. It cannot be undone.</p>
      <div style="max-height:180px;overflow-y:auto;margin-bottom:8px">
        ${sel.map(s => `<div style="padding:7px 0;border-bottom:1px solid var(--stroke);font-size:13px">
          ${esc(s.Name)} <span style="color:var(--faint)">· ${esc(s.UID)}</span></div>`).join('')}
      </div>
      <div class="acts">
        <button class="btn" onclick="closeModal()">Cancel</button>
        <button class="btn bad" id="rmYes">Remove</button>
      </div>`);
    $('#rmYes').onclick = async () => {
      try {
        for (const s of sel) await api('/students/' + encodeURIComponent(s.UID), { method:'DELETE' });
        toast('Removed', `${sel.length} student(s)`, 'ok', 'negative');
        S.sel.clear(); closeModal(); refresh();
      } catch (e) { toast('Could not remove', e.message, 'bad'); }
    };
  };
  $('#btnMark').onclick = async () => {
    const sel = selectedStudents();
    if (!sel.length) return toast('Nothing selected', 'Click a row first', 'warn');
    let done = 0; const skipped = [];
    for (const s of sel) {
      try {
        await api('/mark', { method:'POST', body:JSON.stringify({ uid:s.UID }) });
        done++;
      } catch (e) {
        skipped.push(s.Name);
      }
    }
    if (done) {
      toast('Attendance marked', `${done} student(s)`, 'ok');
      confetti(28); beep([700, 940]);
    }
    if (skipped.length) {
      toast('Already present',
            skipped.length === 1 ? skipped[0] : `${skipped.length} already marked today`,
            'warn');
    }
    refresh();
  };
  $('#btnExport').onclick = doExport;
  $('#btnImport').onclick = () => {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.xlsx,.csv';
    inp.onchange = async () => {
      const fd = new FormData(); fd.append('file', inp.files[0]);
      try {
        const r = await fetch('/api/import', { method:'POST', body:fd });
        const d = await r.json();
        if (!r.ok) throw new Error(d.error);
        toast('Imported', `${d.added} student(s) added`, 'ok', 'export');
        refresh();
      } catch (e) { toast('Import failed', e.message, 'bad'); }
    };
    inp.click();
  };

  // bot
  $('#botSend').onclick = () => ask($('#botInput').value);
  $('#botInput').onkeydown = e => { if (e.key === 'Enter') ask($('#botInput').value); };
  $$('.chip').forEach(c => c.onclick = () => ask(c.dataset.q));
  bubble("Hey! I'm reading your live roster, so every number I give you is real.\\n\\n" +
         'Ask me things like "who is absent", "tell me about Aarav", or\\n' +
         '"house standings" — or just tap one of the chips below.', 'bot');
  // detect Ollama on load so the badge is right before the first question
  api('/ollama').then(d => setBotMode(d.up ? 'AI mode' : 'offline')).catch(() => {});

  // notes
  $('#btnNewNote').onclick = () => {
    S.note = null;
    $('#noteTitle').textContent = 'Untitled note';
    $('#noteBody').value = '';
    $('#noteBody').focus();
  };
  $('#btnSaveNote').onclick = saveNote;
  $('#btnDelNote').onclick = async () => {
    if (!S.note) return toast('No note selected', '', 'warn');
    try {
      await api('/notes/' + encodeURIComponent(S.note), { method:'DELETE' });
      toast('Note deleted', '', 'ok', 'negative');
      S.note = null; $('#noteBody').value = ''; $('#noteTitle').textContent = 'Untitled note';
      loadNotes();
    } catch (e) { toast('Delete failed', e.message, 'bad'); }
  };
  // Autosave removed — it saved half-finished notes and created clutter.
  // Ctrl+S or the Save button now, and a warning if you navigate away dirty.
  let noteDirty = false;
  $('#noteBody').oninput = () => {
    noteDirty = true;
    const t = $('#noteTitle');
    if (t && !t.textContent.endsWith(' •')) t.textContent += ' •';
  };
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      if (S.page === 'notes' && (!window.keyboardOwner ||
          window.keyboardOwner() === 'app')) { saveNote(); noteDirty = false; }
    }
  });
  window.__noteDirty = () => noteDirty;
  window.__noteClean = () => { noteDirty = false; };

  // merged 3.2 panes
  $('#analyticsTab').onclick = () => showAnalyticsPane('analytics');
  $('#housesTab').onclick = () => showAnalyticsPane('houses');
  $('#notesTab').onclick = () => showNotepadPane('notes');
  $('#boardTab').onclick = () => showNotepadPane('whiteboard');

  // console — a real command prompt, not an editable log
  $('#runCmd').onclick = () => runConsoleCommand();
  $('#cmdInput').onkeydown = e => {
    if (e.key === 'Enter') { e.preventDefault(); runConsoleCommand(); }
  };
  $('#btnClearLog').onclick = () => {
    $('#consoleLog').innerHTML = '';
    logLine(now(), 'console buffer cleared', 'w');
  };

  // settings toggles
  const tog = (id, key, fn) => {
    const el = $(id);
    const saved = localStorage.getItem('fa.' + key);
    if (saved !== null) el.classList.toggle('on', saved === '1');
    fn(el.classList.contains('on'));
    el.onclick = () => {
      el.classList.toggle('on');
      const on = el.classList.contains('on');
      localStorage.setItem('fa.' + key, on ? '1' : '0');
      fn(on);
    };
  };
  tog('#swTheme', 'light', on => document.body.classList.toggle('light', on));
  tog('#swMotion', 'reduce', on => document.body.classList.toggle('reduce-motion', on));
  tog('#swSound', 'sound', on => S.fx.sound = on);
  tog('#swConfetti', 'confetti', on => S.fx.confetti = on);

  $('#setThreshold').oninput = e => $('#thVal').textContent = e.target.value;
  $('#btnSaveSettings').onclick = async () => {
    try {
      await api('/settings', { method:'POST', body:JSON.stringify({
        late_time: $('#setLate').value, school_name: $('#setSchool').value,
        threshold: +$('#setThreshold').value, xp_ontime: +$('#setXpOn').value,
        xp_late: +$('#setXpLate').value,
        confirm_frames: +($('#setConfirm')?.value ?? 2),
        match_margin: +($('#setMargin')?.value ?? 0), smtp_user: $('#setSmtpUser').value,
        smtp_pass: $('#setSmtpPass').value, smtp_host: $('#setSmtpHost').value,
        smtp_port: +$('#setSmtpPort').value,
        camera_index: +($('#setCamIndex')?.value ?? 0),
        absent_notify_time: $('#setNotifyTime')?.value || '10:00',
        cert_min_pct: +($('#certPct')?.value ?? 90)
      })});
      toast('Settings saved', '', 'ok');
      $('#setSmtpPass').value = '';
      api('/state').then(d => {
        if ($('#pwSaved'))
          $('#pwSaved').textContent = d.smtp_pass_saved ? '— saved ✓' : '';
        if (d.smtp_pass_saved) $('#setSmtpPass').placeholder = '•••••••• (saved)';
      }).catch(() => {});
    } catch (e) { toast('Could not save', e.message, 'bad'); }
  };
  $('#btnBackup').onclick = doBackup;

  // mail
  $('#mailStudent').onchange = e => {
    const s = S.students.find(x => String(x.UID) === e.target.value);
    if (s) $('#mailTo').value = s['Parent Email'] || '';
  };
  $('#mailTemplate').onclick = () => {
    const s = S.students.find(x => String(x.UID) === $('#mailStudent').value);
    const n = s ? s.Name : 'your child';
    $('#mailBody').value =
      `Dear Parent,\\n\\nThis is a note from ${S.settings.school_name || 'school'} regarding ${n}.\\n\\n` +
      `Attendance to date: ${s ? s['No. of Days Present'] || 0 : 0} days.\\n\\n` +
      `Kind regards,\\nThe class teacher`;
  };
  $('#mailSend').onclick = async () => {
    const to = $('#mailTo').value.trim();
    if (!to) return toast('No recipient', 'Pick a student or type an email', 'warn');
    try {
      await api('/mail', { method:'POST', body:JSON.stringify({
        to, subject:$('#mailSubject').value, body:$('#mailBody').value })});
      toast('Email sent', to, 'ok');
    } catch (e) { toast('Send failed', e.message, 'bad'); }
  };

  // modal veil
  $('#veil').onclick = e => { if (e.target.id === 'veil') closeModal(); };

  // keyboard
  // ── who owns the keyboard right now ────────────────────────────
  // Games use WASD, so the app-wide shortcuts (S = scan, A = add student,
  // 1-9 = switch tab) were firing mid-game: holding A opened the add-student
  // form and S toggled the scanner on and off. Whatever is on top of the
  // screen gets the keys; the global shortcuts only apply when nothing is.
  function keyboardOwner() {
    if (window.__gameActive) return 'arcade';
    const el = document.getElementById('arcadeVeil');
    if (el && el.classList.contains('on')) return 'arcade';
    const veil = document.getElementById('veil');
    if (veil && veil.classList.contains('on')) return 'modal';
    const cmd = document.getElementById('cmdk');
    if (cmd && cmd.classList.contains('on')) return 'palette';
    const gate = document.getElementById('gate');
    if (gate && !gate.classList.contains('hidden')) return 'gate';
    return 'app';
  }
  window.keyboardOwner = keyboardOwner;

  document.addEventListener('keydown', e => {
    // Stop the browser's own Ctrl+K / Cmd+K quick-search from stealing focus
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); return; }
    if (e.key === 'Escape') { closeModal(); closeArcade(); }
    if (e.target.matches('input,textarea,[contenteditable]')) return;

    // A game, dialog, palette or the sign-in screen is holding the keyboard.
    if (keyboardOwner() !== 'app') return;

    if (e.key === '/') { e.preventDefault(); go('home'); $('#q').focus(); }
    if (e.key.toLowerCase() === 's' && !e.ctrlKey) toggleScan();
    if (e.key.toLowerCase() === 'a' && !e.ctrlKey) studentForm();
    if (e.key >= '1' && e.key <= '9') {
      const b = $$('.navbtn')[+e.key - 1]; if (b) go(b.dataset.page);
    }
  });

  // konami easter egg
  const KONAMI = ['ArrowUp','ArrowUp','ArrowDown','ArrowDown','ArrowLeft','ArrowRight','ArrowLeft','ArrowRight','b','a'];
  let kIdx = 0;
  document.addEventListener('keydown', e => {
    if (window.keyboardOwner && window.keyboardOwner() === 'arcade') return;
    kIdx = (e.key === KONAMI[kIdx] || e.key.toLowerCase() === KONAMI[kIdx]) ? kIdx + 1 : 0;
    if (kIdx === KONAMI.length) {
      kIdx = 0;
      document.body.classList.toggle('party');
      const on = document.body.classList.contains('party');
      toast(on ? '🎉 PARTY MODE' : 'Party mode off',
            on ? 'You found it. Nice.' : 'Back to business', 'xp');
      if (on) { confetti(220); beep([523,659,784,1047], .16); }
    }
  });

  // startup chime — plays sound_startup.mp3 if present beside the script,
  // else falls back to a synthesized jingle. refresh() first so S.sounds
  // is populated before we decide which one to use.
  refresh().then(() => playSound('startup', [392, 523, 659, 784], .16));
  setInterval(() => { if (!S.scanning) refresh(); }, 12000);
  logLine(now(), 'FaceAttend 3.2 ready · local runtime connected', 'i');

  api('/state').then(d => {
    $('#dataPath').textContent = d.data_dir || '';
    if ($('#verLabel')) $('#verLabel').textContent = d.version || '';
    const s = d.settings || {};
    $('#setLate').value = s.late_time || '08:30';
    $('#setSchool').value = s.school_name || '';
    $('#setThreshold').value = s.threshold ?? 0.45;
    $('#thVal').textContent = s.threshold ?? 0.45;
    $('#setXpOn').value = s.xp_ontime ?? 5;
    $('#setXpLate').value = s.xp_late ?? 2;
    if ($('#setConfirm')) $('#setConfirm').value = s.confirm_frames ?? 2;
    if ($('#setMargin')) $('#setMargin').value = s.match_margin ?? 0;
    $('#setSmtpUser').value = s.smtp_user || '';
    if ($('#pwSaved'))
      $('#pwSaved').textContent = d.smtp_pass_saved ? '— saved ✓' : '';
    if (d.smtp_pass_saved && $('#setSmtpPass'))
      $('#setSmtpPass').placeholder = '•••••••• (saved)';
    $('#setSmtpHost').value = s.smtp_host || 'smtp.gmail.com';
    $('#setSmtpPort').value = s.smtp_port || 587;
    if ($('#setCamIndex')) $('#setCamIndex').value = s.camera_index ?? 0;
    if ($('#setNotifyTime')) $('#setNotifyTime').value = s.absent_notify_time || '10:00';
    if ($('#certPct')) $('#certPct').value = s.cert_min_pct ?? 90;
  }).catch(() => {});

  // 3.2 feature wiring — badges, certificates, calendar, QR, audit,
  // bulk import, section moves, absent notifications, vault.
  initExtras();
  initDay();
  initDev();
  initGate();
}


/* ══════════════════════════════════════════════════════════════════
   3.2 FEATURE SET — badges, certificates, calendar, QR, audit,
   bulk import, section moves, absent notifications, voice, vault
   ══════════════════════════════════════════════════════════════════ */

/* ── voice feedback (browser speech synthesis, no extra install) ── */
function speak(text) {
  if (!S.fx.voice || !window.speechSynthesis) return;
  try {
    const u = new SpeechSynthesisUtterance(text);
    u.rate = 1.05; u.pitch = 1.0; u.volume = 0.9;
    speechSynthesis.cancel();
    speechSynthesis.speak(u);
  } catch (e) {}
}

/* ── badges ─────────────────────────────────────────────────────── */
async function openBadges() {
  try {
    const d = await api('/badges');
    const withAny = d.students.filter(s => s.earned > 0);
    modal(`
      <h2>🏅 Badges</h2>
      <p class="sub">${withAny.length} of ${d.students.length} students have earned something.</p>
      <div style="max-height:58vh;overflow-y:auto;display:flex;flex-direction:column;gap:12px">
        ${d.students.map(s => `
          <div class="card" style="padding:13px">
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:9px">
              <b style="font-size:13.5px">${esc(s.name)}</b>
              <span style="color:var(--faint);font-size:11.5px">${esc(s.grade||'')} · ${esc(s.house||'')}</span>
              <span class="pill xp" style="margin-left:auto">${s.earned}/6</span>
            </div>
            <div style="display:flex;gap:7px;flex-wrap:wrap">
              ${s.badges.map(b => `
                <div title="${esc(b.desc)} — ${esc(b.detail)}"
                     style="flex:1;min-width:88px;padding:9px;border-radius:11px;text-align:center;
                            background:${b.earned ? 'rgba(255,209,102,.13)' : 'var(--surface-2)'};
                            border:1px solid ${b.earned ? 'var(--xp)' : 'var(--stroke)'};
                            opacity:${b.earned ? 1 : .5}">
                  <div style="font-size:20px">${b.icon}</div>
                  <div style="font-size:10.5px;font-weight:700;margin-top:3px">${esc(b.name)}</div>
                  <div style="font-size:9.5px;color:var(--faint);margin-top:2px">${esc(b.detail)}</div>
                  <div style="height:3px;border-radius:2px;background:var(--stroke);margin-top:5px;overflow:hidden">
                    <div style="height:100%;width:${b.progress}%;background:var(--grad)"></div>
                  </div>
                </div>`).join('')}
            </div>
          </div>`).join('')}
      </div>
      <div class="acts"><button class="btn" onclick="closeModal()">Close</button></div>`);
  } catch (e) { toast('Could not load badges', e.message, 'bad'); }
}

/* ── certificates ───────────────────────────────────────────────── */
async function makeCertificates() {
  const pct = parseFloat(document.getElementById('certPct').value) || 90;
  toast('Generating…', 'Building PDFs', 'info');
  try {
    const d = await api('/certificates', { method:'POST', body:JSON.stringify({ min_pct: pct }) });
    modal(`
      <h2>📜 Certificates</h2>
      <p class="sub">${d.made.length} generated · ${d.skipped.length} below ${pct}%</p>
      ${d.made.length ? `<div style="max-height:40vh;overflow-y:auto">
        ${d.made.map(m => `<div style="display:flex;justify-content:space-between;padding:9px 0;
             border-bottom:1px solid var(--stroke);font-size:13px">
             <span>${esc(m.name)}</span>
             <a href="/api/certificates/${encodeURIComponent(m.file)}" target="_blank"
                style="color:var(--accent);text-decoration:none">${m.pct}% · open ↗</a></div>`).join('')}
        </div>` : '<p style="color:var(--muted);font-size:13px">Nobody met the threshold.</p>'}
      <p style="font-size:11.5px;color:var(--faint);margin-top:12px">Saved to: ${esc(d.folder)}</p>
      <div class="acts"><button class="btn" onclick="closeModal()">Close</button></div>`);
    if (d.made.length) { confetti(60); beep([660,880,1100]); }
  } catch (e) { toast('Certificate failed', e.message, 'bad'); }
}

/* ── calendar heatmap ───────────────────────────────────────────── */
let calY = new Date().getFullYear(), calM = new Date().getMonth() + 1;
async function openCalendar() {
  try {
    const d = await api(`/calendar?year=${calY}&month=${calM}`);
    const cells = [];
    for (let i = 0; i < d.first_weekday; i++) cells.push('<div></div>');
    d.days.forEach(day => {
      const a = day.pct / 100;
      const bg = day.count === 0
        ? (day.weekend ? 'var(--surface)' : 'var(--surface-2)')
        : `rgba(46,230,168,${0.18 + a * 0.72})`;
      cells.push(`<div title="${day.count} present (${day.pct.toFixed(0)}%)"
        style="aspect-ratio:1;border-radius:9px;background:${bg};
               border:1px solid var(--stroke);display:grid;place-items:center;
               font-size:11.5px;font-weight:650;
               color:${day.count ? '#04241a' : 'var(--faint)'}">${day.day}</div>`);
    });
    modal(`
      <h2>📅 ${d.month_name} ${d.year}</h2>
      <p class="sub">Darker green means more students present that day.</p>
      <div style="display:flex;gap:8px;margin-bottom:12px">
        <button class="btn sm" id="calPrev">← Prev</button>
        <button class="btn sm" id="calNext">Next →</button>
        <button class="btn sm" id="calToday" style="margin-left:auto">Today</button>
      </div>
      <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:5px;margin-bottom:6px">
        ${['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].map(x =>
          `<div style="text-align:center;font-size:10px;font-weight:700;color:var(--faint)">${x}</div>`).join('')}
      </div>
      <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:5px">${cells.join('')}</div>
      <div class="acts"><button class="btn" onclick="closeModal()">Close</button></div>`);
    document.getElementById('calPrev').onclick = () => {
      calM--; if (calM < 1) { calM = 12; calY--; } openCalendar(); };
    document.getElementById('calNext').onclick = () => {
      calM++; if (calM > 12) { calM = 1; calY++; } openCalendar(); };
    document.getElementById('calToday').onclick = () => {
      const n = new Date(); calY = n.getFullYear(); calM = n.getMonth() + 1; openCalendar(); };
  } catch (e) { toast('Calendar failed', e.message, 'bad'); }
}

/* ── QR pass ────────────────────────────────────────────────────── */
function showQr(uid) {
  const s = S.students.find(x => String(x.UID) === String(uid));
  if (!s) return toast('Pick a student first', '', 'warn');
  modal(`
    <h2>🔳 ${esc(s.Name)}</h2>
    <p class="sub">Scannable identity pass — UID ${esc(s.UID)}</p>
    <div style="display:grid;place-items:center;padding:16px;background:#fff;border-radius:16px">
      <img src="/api/qr/${encodeURIComponent(uid)}" style="width:240px;height:240px"
           onerror="this.replaceWith(Object.assign(document.createElement('p'),
                    {textContent:'QR needs the qrcode package: pip install qrcode',
                     style:'color:#333;font-size:13px;padding:30px;text-align:center'}))">
    </div>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Close</button>
      <a class="btn primary" href="/api/qr/${encodeURIComponent(uid)}" download
         style="text-decoration:none">Download PNG</a>
    </div>`);
}

/* ── audit log ──────────────────────────────────────────────────── */
async function openAudit() {
  try {
    const d = await api('/audit');
    modal(`
      <h2>🧾 Audit log</h2>
      <p class="sub">${d.total} total entries · newest first</p>
      <div style="max-height:56vh;overflow-y:auto">
        ${d.rows.length ? `<table style="width:100%;font-size:12.5px">
          <thead><tr>${['Date','Time','Name','Status','XP'].map(h =>
            `<th style="text-align:left;padding:7px;color:var(--muted);font-size:10.5px">${h}</th>`).join('')}</tr></thead>
          <tbody>${d.rows.map(r => `<tr style="border-bottom:1px solid var(--stroke)">
            <td style="padding:7px">${esc(r.date)}</td>
            <td style="padding:7px;font-family:var(--mono)">${esc(r.time)}</td>
            <td style="padding:7px">${esc(r.name)}</td>
            <td style="padding:7px"><span class="pill ${r.status==='late'?'late':'on'}">${esc(r.status)}</span></td>
            <td style="padding:7px;color:var(--xp)">+${esc(r.xp)}</td></tr>`).join('')}
          </tbody></table>`
          : '<p style="color:var(--muted);font-size:13px">Nothing logged yet.</p>'}
      </div>
      <div class="acts"><button class="btn" onclick="closeModal()">Close</button></div>`);
  } catch (e) { toast('Audit failed', e.message, 'bad'); }
}

/* ── bulk import ────────────────────────────────────────────────── */
function openBulk() {
  modal(`
    <h2>📋 Bulk add students</h2>
    <p class="sub">One student per line: Name, Roll, Grade, House, Phone, Email</p>
    <div class="field">
      <textarea id="bulkText" rows="10" placeholder="Aarav Arjun, 8C02, 8C, Orange, 9876543210, parent@example.com
Priya Nair, 8C07, 8C, Green
Rohan Das, 9A11, 9A, Blue"></textarea>
    </div>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="bulkGo">Add them</button>
    </div>`);
  document.getElementById('bulkGo').onclick = async () => {
    const text = document.getElementById('bulkText').value;
    if (!text.trim()) return toast('Nothing to add', '', 'warn');
    try {
      const d = await api('/bulk', { method:'POST', body:JSON.stringify({ text }) });
      toast('Bulk import done', `${d.added} student(s) added`, 'ok');
      if (d.failed.length) toast('Some rows failed', d.failed[0], 'warn');
      closeModal(); refresh(); confetti(40);
    } catch (e) { toast('Import failed', e.message, 'bad'); }
  };
}

/* ── move students between grade / house ────────────────────────── */
function openSection() {
  const sel = selectedStudents();
  if (!sel.length) return toast('Select students first', 'Click one or more rows', 'warn');
  modal(`
    <h2>🔀 Move ${sel.length} student${sel.length>1?'s':''}</h2>
    <p class="sub">Leave a field blank to keep it unchanged.</p>
    <div class="grid2">
      <div class="field"><label>New grade</label><input id="secGrade" placeholder="e.g. 9A"></div>
      <div class="field"><label>New house</label><select id="secHouse">
        <option value="">— keep —</option><option>Red</option><option>Blue</option>
        <option>Green</option><option>Orange</option></select></div>
    </div>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="secGo">Apply</button>
    </div>`);
  document.getElementById('secGo').onclick = async () => {
    try {
      const d = await api('/section', { method:'POST', body:JSON.stringify({
        uids: sel.map(s => String(s.UID)),
        grade: document.getElementById('secGrade').value.trim(),
        house: document.getElementById('secHouse').value })});
      toast('Moved', `${d.changed} student(s) updated`, 'ok');
      closeModal(); refresh();
    } catch (e) { toast('Move failed', e.message, 'bad'); }
  };
}

/* ── notify absent parents ──────────────────────────────────────── */
async function notifyAbsent() {
  const absent = S.students.filter(s => !S.present.includes(String(s.UID)));
  if (!absent.length) return toast('Nobody is absent', 'Nothing to send', 'info');
  modal(`
    <h2>📨 Notify absent parents</h2>
    <p class="sub">${absent.length} student(s) are not marked present today.</p>
    <div style="max-height:34vh;overflow-y:auto;margin-bottom:8px">
      ${absent.map(s => `<div style="display:flex;justify-content:space-between;padding:7px 0;
          border-bottom:1px solid var(--stroke);font-size:13px">
          <span>${esc(s.Name)}</span>
          <span style="color:${s['Parent Email'] ? 'var(--muted)' : 'var(--bad)'};font-size:12px">
            ${esc(s['Parent Email'] || 'no email')}</span></div>`).join('')}
    </div>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="notifyGo">Send emails</button>
    </div>`);
  document.getElementById('notifyGo').onclick = async () => {
    toast('Sending…', 'Contacting mail server', 'info');
    try {
      const d = await api('/notify-absent', { method:'POST', body:JSON.stringify({}) });
      toast('Emails sent', `${d.sent.length} sent, ${d.skipped.length} skipped`, 'ok');
      closeModal();
    } catch (e) { toast('Send failed', e.message, 'bad'); }
  };
}

/* ── awards panel wiring ────────────────────────────────────────── */
async function refreshAwards() {
  try {
    const d = await api('/badges');
    const awarded = d.students.reduce((n, s) => n + s.earned, 0);
    document.getElementById('bdgAwarded').textContent = awarded;
    document.getElementById('bdgStudents').textContent =
      d.students.filter(s => s.earned > 0).length;
    document.getElementById('bdgTotal').textContent = Object.keys(d.defs).length;
    const sel = document.getElementById('qrStudent');
    sel.innerHTML = '<option value="">Select a student…</option>' +
      S.students.map(s => `<option value="${esc(s.UID)}">${esc(s.Name)}</option>`).join('');
  } catch (e) {}
}

function initExtras() {
  const on = (id, fn) => { const el = document.getElementById(id); if (el) el.onclick = fn; };
  on('btnSignOut', signOut);
  on('btnSwitch', switchAccount);
  on('btnChangePw', () => {
    modal(`
      <h2>🔑 Change password</h2>
      <p class="sub">You will stay signed in on this device.</p>
      <div class="field"><label>Current password</label>
        <input id="pwOld" type="password"></div>
      <div class="field"><label>New password</label>
        <input id="pwNew" type="password"></div>
      <div class="field"><label>Confirm new password</label>
        <input id="pwNew2" type="password"></div>
      <div class="acts">
        <button class="btn" onclick="closeModal()">Cancel</button>
        <button class="btn primary" id="pwGo">Update</button>
      </div>`);
    document.getElementById('pwGo').onclick = async () => {
      const a = document.getElementById('pwNew').value;
      if (a !== document.getElementById('pwNew2').value)
        return toast('Passwords do not match', '', 'warn');
      try {
        await api('/auth/password', { method:'POST', body:JSON.stringify({
          old: document.getElementById('pwOld').value, new: a })});
        toast('Password updated', '', 'ok'); closeModal();
      } catch (e) { toast('Could not update', e.message, 'bad'); }
    };
  });
  on('btnUnmark', unmarkStudents);
  on('btnAbsence', recordAbsence);
  on('btnReport', openReport);
  on('btnVerify', verifyLog);
  on('btnSchedule', openSchedule);
  on('btnCloseRegister', closeRegister);
  on('btnCameras', openCameras);
  on('btnRecogTest', testRecognition);
  on('btnBadges', openBadges);
  on('btnCerts', makeCertificates);
  on('btnCalendar', openCalendar);
  on('btnAudit', openAudit);
  on('btnBulk', openBulk);
  on('btnSection', openSection);
  on('btnNotify', notifyAbsent);
  on('btnQr', () => showQr(document.getElementById('qrStudent').value));
  on('btnEncrypt', async () => {
    try {
      const d = await api('/vault', { method:'POST', body:JSON.stringify({ action:'encrypt' }) });
      toast(d.ok ? 'Roster encrypted' : 'Encryption failed', d.file || '', d.ok ? 'ok' : 'bad');
    } catch (e) { toast('Encryption failed', e.message, 'bad'); }
  });

  // extra setting toggles
  const tog = (id, key, fn) => {
    const el = document.getElementById(id); if (!el) return;
    const saved = localStorage.getItem('fa.' + key);
    if (saved !== null) el.classList.toggle('on', saved === '1');
    fn(el.classList.contains('on'));
    el.onclick = () => {
      el.classList.toggle('on');
      const v = el.classList.contains('on');
      localStorage.setItem('fa.' + key, v ? '1' : '0');
      fn(v);
    };
  };
  tog('swVoice', 'voice', v => S.fx.voice = v);
  tog('swLiveness', 'liveness', v => api('/settings', {
    method:'POST', body:JSON.stringify({ liveness_check: v }) }).catch(()=>{}));
  tog('swContrast', 'contrast', v => document.body.classList.toggle('contrast', v));
  tog('swEvidence', 'evidence', v => api('/settings', {
    method:'POST', body:JSON.stringify({ keep_scan_photos: v }) }).catch(()=>{}));
  tog('swAutoBackup', 'autobk', v => api('/settings', {
    method:'POST', body:JSON.stringify({ auto_backup: v }) }).catch(()=>{}));
  tog('swEnforce', 'enforce', v => api('/schedule', {
    method:'POST', body:JSON.stringify({ enforce: v }) })
      .then(r => paintWindow(r.now)).catch(()=>{}));
  tog('swApi', 'apisrv', v => api('/settings', {
    method:'POST', body:JSON.stringify({ api_server: v }) }).catch(()=>{}));
  tog('swAutoNotify', 'autonotify', v => api('/settings', {
    method:'POST', body:JSON.stringify({ auto_notify_absent: v }) }).catch(()=>{}));
}


/* ══════════════════════════════════════════════════════════════════
   SIGN IN / SIGN UP
   The app stays hidden behind this gate until someone is signed in.
   ══════════════════════════════════════════════════════════════════ */
let gateMode = 'login';          // or 'register'
let currentUser = null;

const gEl = id => document.getElementById(id);

function gateError(msg) {
  const e = gEl('gateErr');
  if (e) e.textContent = msg || '';
}

function setGateMode(mode) {
  gateMode = mode;
  const isReg = mode === 'register';
  gEl('gateSub').textContent = isReg
    ? 'Create your account to get started'
    : 'Sign in to continue';
  gEl('gateGo').textContent = isReg ? 'Create account' : 'Sign in';
  gEl('gateDisplayWrap').style.display = isReg ? '' : 'none';
  gEl('gateConfirmWrap').style.display = isReg ? '' : 'none';
  gEl('gateMeter').style.display = isReg ? '' : 'none';
  gEl('gateSwapText').textContent = isReg ? 'Already have an account?' : 'First time here?';
  gEl('gateSwap').textContent = isReg ? 'Sign in instead' : 'Create an account';
  gEl('gatePass').setAttribute('autocomplete', isReg ? 'new-password' : 'current-password');
  gateError('');
}

function passwordScore(pw) {
  let n = 0;
  if (pw.length >= 6) n++;
  if (pw.length >= 10) n++;
  if (/[A-Z]/.test(pw) && /[a-z]/.test(pw)) n++;
  if (/[0-9]/.test(pw)) n++;
  if (/[^A-Za-z0-9]/.test(pw)) n++;
  return Math.min(4, n);
}

function paintMeter() {
  const bar = gEl('gateMeter')?.querySelector('i');
  if (!bar) return;
  const n = passwordScore(gEl('gatePass').value);
  const cols = ['var(--bad)', 'var(--bad)', 'var(--warn)', 'var(--ok)', 'var(--ok)'];
  bar.style.width = `${n * 25}%`;
  bar.style.background = cols[n];
}

async function gateSubmit(e) {
  if (e) e.preventDefault();
  const btn = gEl('gateGo');
  const username = gEl('gateUser').value.trim();
  const password = gEl('gatePass').value;
  if (!username || !password) return gateError('Fill in both fields.');

  if (gateMode === 'register') {
    if (password !== gEl('gateConfirm').value)
      return gateError('The two passwords do not match.');
    if (password.length < 6)
      return gateError('Password must be at least 6 characters.');
  }

  btn.disabled = true;
  btn.textContent = gateMode === 'register' ? 'Creating…' : 'Signing in…';
  gateError('');
  try {
    const body = { username, password };
    if (gateMode === 'register') body.display = gEl('gateDisplay').value.trim();
    const d = await api('/auth/' + (gateMode === 'register' ? 'register' : 'login'),
                        { method:'POST', body:JSON.stringify(body) });
    if (d.token) localStorage.setItem('fa.token', d.token);
    gEl('gatePass').value = '';
    if (gEl('gateConfirm')) gEl('gateConfirm').value = '';
    await enterApp();
  } catch (err) {
    gateError(err.message || 'Could not sign in.');
    gEl('gatePass').focus();
    if (typeof beep === 'function') beep([300, 220], .1);
  }
  btn.disabled = false;
  btn.textContent = gateMode === 'register' ? 'Create account' : 'Sign in';
}

async function enterApp() {
  const st = await api('/auth/status');
  currentUser = st.user;
  gEl('gate').classList.add('hidden');
  const who = currentUser?.display || currentUser?.username || '';
  toast(`Welcome back, ${who}`, '', 'ok');
  if (typeof confetti === 'function') confetti(28);
  const badge = gEl('whoami');
  if (badge) badge.textContent = who;
  refresh();
}

async function signOut() {
  try { await api('/auth/logout', { method:'POST', body:'{}' }); } catch (e) {}
  localStorage.removeItem('fa.token');
  currentUser = null;
  gEl('gate').classList.remove('hidden');
  try {
    const d = await api('/auth/profiles');
    window.__profiles = d.profiles || [];
    if (window.__profiles.length > 1) { showPicker(window.__profiles); return; }
  } catch (e) {}
  showFullForm('login');
}


/* ── account picker ─────────────────────────────────────────────── */
const AV_COLOURS = ['#5b8cff','#a56bff','#22d3ee','#2ee6a8','#ffb020','#ff5d73'];

function avatarFor(name, username) {
  const initials = String(name || username || '?')
    .split(/[ ]+/).map(w => w[0]).slice(0, 2).join('').toUpperCase();
  let h = 0;
  for (const ch of String(username || '')) h = (h * 31 + ch.charCodeAt(0)) % 997;
  return { initials, colour: AV_COLOURS[h % AV_COLOURS.length] };
}

let pickedUser = null;

function showPicker(profiles) {
  const grid = gEl('pickGrid');
  grid.innerHTML = profiles.map(pr => {
    const a = avatarFor(pr.display, pr.username);
    return `<button type="button" class="picktile" data-u="${pr.username}">
      <span class="pav" style="background:${a.colour}">${a.initials}</span>
      <span class="pinfo">
        <span class="pname">${pr.display || pr.username}</span>
        <span class="prole">${pr.role === 'admin' ? 'Administrator' : 'Teacher'}
          · @${pr.username}</span>
      </span>
      <span class="parrow">›</span>
    </button>`;
  }).join('');

  grid.querySelectorAll('.picktile').forEach(t => t.onclick = () => {
    const pr = profiles.find(x => x.username === t.dataset.u);
    choosePicked(pr);
  });

  gEl('gatePicker').style.display = '';
  gEl('gateForm').style.display = 'none';
  gEl('gateSub').textContent = 'Choose an account to continue';
  const swap = document.querySelector('.gateswap');
  if (swap) swap.style.display = 'none';
}

function choosePicked(pr) {
  pickedUser = pr;
  gEl('gatePicker').style.display = 'none';
  gEl('gateForm').style.display = '';
  gEl('pickBack').style.display = '';
  setGateMode('login');

  // Show who you are signing in as, so only the password is left to type.
  const a = avatarFor(pr.display, pr.username);
  const old = document.getElementById('gateWho');
  if (old) old.remove();
  const who = document.createElement('div');
  who.className = 'gatewho';
  who.id = 'gateWho';
  who.innerHTML = `<span class="pav" style="background:${a.colour}">${a.initials}</span>
    <span><b style="font-size:13px">${pr.display || pr.username}</b>
    <div style="font-size:11px;color:var(--faint)">@${pr.username}</div></span>`;
  gEl('gateForm').insertBefore(who, gEl('pickBack').nextSibling);

  gEl('gateUser').value = pr.username;
  gEl('gateUser').closest('.gatefield').style.display = 'none';
  gEl('gateSub').textContent = 'Enter your password';
  const swap = document.querySelector('.gateswap');
  if (swap) swap.style.display = 'none';
  setTimeout(() => gEl('gatePass').focus(), 60);
}

function showFullForm(mode) {
  pickedUser = null;
  const who = document.getElementById('gateWho');
  if (who) who.remove();
  gEl('gatePicker').style.display = 'none';
  gEl('gateForm').style.display = '';
  gEl('pickBack').style.display = window.__profiles?.length ? '' : 'none';
  gEl('gateUser').closest('.gatefield').style.display = '';
  gEl('gateUser').value = '';
  gEl('gatePass').value = '';
  const swap = document.querySelector('.gateswap');
  if (swap) swap.style.display = '';
  setGateMode(mode || 'login');
  gEl('gateUser').focus();
}

async function backToPicker() {
  const who = document.getElementById('gateWho');
  if (who) who.remove();
  gEl('gatePass').value = '';
  gateError('');
  gEl('gateUser').closest('.gatefield').style.display = '';
  gEl('pickBack').style.display = 'none';
  if (window.__profiles?.length) showPicker(window.__profiles);
  else showFullForm('login');
}

async function switchAccount() {
  try { await api('/auth/logout', { method:'POST', body:'{}' }); } catch (e) {}
  localStorage.removeItem('fa.token');
  currentUser = null;
  gEl('gate').classList.remove('hidden');
  try {
    const d = await api('/auth/profiles');
    window.__profiles = d.profiles || [];
    if (window.__profiles.length) { showPicker(window.__profiles); return; }
  } catch (e) {}
  showFullForm('login');
}

async function initGate() {
  gEl('gateForm').onsubmit = gateSubmit;
  gEl('gateSwap').onclick = () =>
    setGateMode(gateMode === 'login' ? 'register' : 'login');
  gEl('gatePass').oninput = paintMeter;
  gEl('gateEye').onclick = () => {
    const f = gEl('gatePass');
    f.type = f.type === 'password' ? 'text' : 'password';
    gEl('gateEye').textContent = f.type === 'password' ? '👁' : '🙈';
  };

  gEl('pickOther').onclick = () => showFullForm('login');
  gEl('pickBack').onclick = backToPicker;

  try {
    const st = await api('/auth/status');
    if (st.signed_in) { await enterApp(); return; }

    if (st.setup_needed) {
      setGateMode('register');
      gEl('gateSub').textContent = 'Create the first account for this school';
      const swap = document.querySelector('.gateswap');
      if (swap) swap.style.display = 'none';
      gEl('gateUser').focus();
      return;
    }

    // More than one account on this machine? Offer them as tiles rather
    // than making everyone type a username.
    const d = await api('/auth/profiles');
    window.__profiles = d.profiles || [];
    if (window.__profiles.length > 1) { showPicker(window.__profiles); return; }
    if (window.__profiles.length === 1) { choosePicked(window.__profiles[0]); return; }
  } catch (e) {}

  setGateMode('login');
  gEl('gateUser').focus();
}


/* ══════════════════════════════════════════════════════════════════
   CORRECTIONS, REPORTS, CAMERA SOURCES, LOG INTEGRITY
   ══════════════════════════════════════════════════════════════════ */
const REASONS = ['Sick','Excused','Late with note','Medical','Family',
                 'Marked in error','Other'];

async function unmarkStudents() {
  const sel = selectedStudents().filter(s => S.present.includes(String(s.UID)));
  if (!sel.length)
    return toast('Nothing to unmark', 'Select students who are marked present', 'warn');
  modal(`
    <h2>↩ Unmark ${sel.length} student${sel.length>1?'s':''}</h2>
    <p class="sub">This reverses the XP, day count and streak from that mark,
       and records the correction in the audit log.</p>
    <div style="max-height:26vh;overflow-y:auto;margin-bottom:12px">
      ${sel.map(s => `<div style="padding:7px 0;border-bottom:1px solid var(--stroke);
        font-size:13px">${esc(s.Name)}</div>`).join('')}
    </div>
    <div class="field"><label>Reason</label>
      <select id="umReason">${REASONS.map(r => `<option>${r}</option>`).join('')}</select>
    </div>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn bad" id="umGo">Unmark</button>
    </div>`);
  document.getElementById('umGo').onclick = async () => {
    const reason = document.getElementById('umReason').value;
    let n = 0;
    for (const st of sel) {
      try { await api('/unmark', { method:'POST',
        body:JSON.stringify({ uid:String(st.UID), reason }) }); n++; }
      catch (e) { toast('Could not unmark ' + st.Name, e.message, 'bad'); }
    }
    toast('Corrected', `${n} student(s) unmarked`, 'ok');
    closeModal(); refresh();
  };
}

async function recordAbsence() {
  const sel = selectedStudents().filter(s => !S.present.includes(String(s.UID)));
  if (!sel.length)
    return toast('Nothing to record', 'Select students who are absent', 'warn');
  modal(`
    <h2>📝 Absence reason</h2>
    <p class="sub">Logs why ${sel.length} student${sel.length>1?'s are':' is'} away.
       It does not mark them present.</p>
    <div class="field"><label>Reason</label>
      <select id="abReason">${REASONS.filter(r => r !== 'Marked in error')
        .map(r => `<option>${r}</option>`).join('')}</select>
    </div>
    <div class="field"><label>Note (optional)</label>
      <input id="abNote" placeholder="e.g. doctor's appointment"></div>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="abGo">Record</button>
    </div>`);
  document.getElementById('abGo').onclick = async () => {
    const note = document.getElementById('abNote').value.trim();
    const reason = document.getElementById('abReason').value + (note ? ` — ${note}` : '');
    let n = 0;
    for (const st of sel) {
      try { await api('/absence', { method:'POST',
        body:JSON.stringify({ uid:String(st.UID), reason }) }); n++; }
      catch (e) {}
    }
    toast('Recorded', `${n} absence reason(s) logged`, 'ok');
    closeModal();
  };
}

async function openReport() {
  try {
    const d = await api('/report');
    const band = p => p >= 90 ? 'on' : p >= 75 ? 'late' : 'off';
    modal(`
      <h2>📈 Attendance report</h2>
      <p class="sub">${d.school_days} school day(s)${d.from ? ` · ${d.from} to ${d.to}` : ''}
         · average ${d.average}%</p>
      <div style="max-height:52vh;overflow-y:auto">
        <table style="width:100%;font-size:12.5px">
          <thead><tr>${['Student','Grade','Days','Late','%'].map(h =>
            `<th style="text-align:left;padding:7px;color:var(--muted);font-size:10.5px">${h}</th>`).join('')}</tr></thead>
          <tbody>${d.students.map(r => `<tr style="border-bottom:1px solid var(--stroke)">
            <td style="padding:8px">${esc(r.name)}
              ${r.absences.length ? `<div style="font-size:10.5px;color:var(--faint)">
                ${r.absences.slice(0,2).map(a => esc(a.reason)).join(', ')}</div>` : ''}</td>
            <td style="padding:8px;color:var(--muted)">${esc(r.grade||'—')}</td>
            <td style="padding:8px">${r.days_present}</td>
            <td style="padding:8px;color:var(--warn)">${r.days_late || '—'}</td>
            <td style="padding:8px"><span class="pill ${band(r.pct)}">${r.pct}%</span></td>
          </tr>`).join('')}</tbody>
        </table>
      </div>
      <div class="acts">
        <button class="btn" onclick="closeModal()">Close</button>
        <button class="btn ok" id="repExport">⬇ Export to Excel</button>
      </div>`);
    document.getElementById('repExport').onclick = async () => {
      try {
        const r = await api('/report/export', { method:'POST', body:'{}' });
        toast('Report exported', r.file, 'ok');
      } catch (e) { toast('Export failed', e.message, 'bad'); }
    };
  } catch (e) { toast('Report failed', e.message, 'bad'); }
}

async function verifyLog() {
  try {
    const d = await api('/audit/verify');
    modal(`
      <h2>${d.ok ? '🔐 Log is intact' : '⚠ Log has been altered'}</h2>
      <p class="sub">${d.checked} entr${d.checked === 1 ? 'y' : 'ies'} checked.</p>
      <div class="card" style="padding:16px;border-color:${d.ok ? 'var(--ok)' : 'var(--bad)'}">
        <p style="font-size:13px;line-height:1.6;color:${d.ok ? 'var(--ok)' : 'var(--bad)'}">
          ${d.ok
            ? 'Every row matches its hash. Nothing has been edited or removed since it was written.'
            : esc(d.problem)}
        </p>
      </div>
      <p style="font-size:11.5px;color:var(--faint);margin-top:12px;line-height:1.6">
        Each row is hashed together with the one before it, so changing any
        past entry breaks the chain from that point onward.
      </p>
      <div class="acts"><button class="btn" onclick="closeModal()">Close</button></div>`);
    if (typeof beep === 'function') beep(d.ok ? [660, 880] : [330, 220], .1);
  } catch (e) { toast('Verify failed', e.message, 'bad'); }
}

async function openCameras() {
  let cams = [];
  try { cams = (await api('/cameras')).cameras || []; } catch (e) {}
  if (!cams.length) cams = [{ name:'Main gate', source:'0' }];

  const rows = () => cams.map((c, i) => `
    <div class="card" style="padding:12px;margin-bottom:8px">
      <div class="grid2" style="gap:8px">
        <div class="field" style="margin:0"><label>Name</label>
          <input data-i="${i}" data-k="name" value="${esc(c.name)}"></div>
        <div class="field" style="margin:0"><label>Source</label>
          <input data-i="${i}" data-k="source" value="${esc(c.source)}"
                 placeholder="0 or rtsp://…"></div>
      </div>
      <div class="row" style="margin-top:8px">
        <button class="btn sm" data-test="${i}">Test</button>
        <button class="btn sm bad" data-del="${i}">Remove</button>
        <span data-res="${i}" style="font-size:11.5px;color:var(--faint);
              align-self:center;margin-left:4px"></span>
      </div>
    </div>`).join('');

  const paint = () => {
    document.getElementById('camList').innerHTML = rows();
    document.querySelectorAll('#camList input').forEach(el => el.oninput = () => {
      cams[+el.dataset.i][el.dataset.k] = el.value;
    });
    document.querySelectorAll('#camList [data-del]').forEach(b => b.onclick = () => {
      cams.splice(+b.dataset.del, 1); paint();
    });
    document.querySelectorAll('#camList [data-test]').forEach(b => b.onclick = async () => {
      const i = +b.dataset.test;
      const out = document.querySelector(`[data-res="${i}"]`);
      out.textContent = 'testing…'; out.style.color = 'var(--muted)';
      try {
        const r = await api('/cameras/test', { method:'POST',
          body:JSON.stringify({ source: cams[i].source }) });
        if (r.ok) { out.textContent = `✓ ${r.width}×${r.height}`; out.style.color = 'var(--ok)'; }
        else { out.textContent = '✕ ' + (r.error || 'failed'); out.style.color = 'var(--bad)'; }
      } catch (e) { out.textContent = '✕ ' + e.message; out.style.color = 'var(--bad)'; }
    });
  };

  modal(`
    <h2>📹 Camera sources</h2>
    <p class="sub">A number is a USB webcam. Anything else is treated as a
       stream URL, so IP cameras at each gate work here.</p>
    <div id="camList" style="max-height:44vh;overflow-y:auto"></div>
    <button class="btn sm" id="camAdd">＋ Add camera</button>
    <p style="font-size:11px;color:var(--faint);margin-top:12px;line-height:1.6">
      Typical URL: <span style="font-family:var(--mono)">rtsp://user:pass@192.168.1.50:554/stream1</span><br>
      Test each one before saving — a wrong password or path fails silently otherwise.
    </p>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="camSave">Save</button>
    </div>`);
  paint();
  document.getElementById('camAdd').onclick = () => {
    cams.push({ name:`Gate ${cams.length + 1}`, source:'' }); paint();
  };
  document.getElementById('camSave').onclick = async () => {
    try {
      const r = await api('/cameras', { method:'POST', body:JSON.stringify({ cameras: cams }) });
      toast('Cameras saved', `${r.cameras.length} source(s)`, 'ok');
      closeModal();
    } catch (e) { toast('Save failed', e.message, 'bad'); }
  };
}


async function testRecognition() {
  toast('Testing…', 'Comparing every enrolled face', 'info');
  try {
    const d = await api('/recognition/test');
    if (!d.ok) return toast('Cannot test yet', d.error, 'warn');
    const good = !d.risky.length && (d.margin === null || d.margin >= 0.08);
    modal(`
      <h2>🎯 Recognition accuracy</h2>
      <p class="sub">${d.students} student(s) · ${d.encodings} face encoding(s)
         · threshold ${d.threshold}</p>

      <div class="card" style="padding:15px;margin-bottom:12px;
           border-color:${good ? 'var(--ok)' : d.risky.length ? 'var(--bad)' : 'var(--warn)'}">
        <b style="font-size:13.5px;color:${good ? 'var(--ok)' : d.risky.length ? 'var(--bad)' : 'var(--warn)'}">
          ${esc(d.verdict)}</b>
      </div>

      <div class="grid2" style="gap:10px;margin-bottom:12px">
        <div class="card" style="padding:13px">
          <div style="font-size:10.5px;color:var(--muted);text-transform:uppercase;
               letter-spacing:.8px;font-weight:750">Same person</div>
          <div style="font-size:21px;font-weight:800;color:var(--ok)">${d.same_avg ?? '—'}</div>
          <div style="font-size:11px;color:var(--faint)">worst ${d.same_worst ?? '—'} · lower is better</div>
        </div>
        <div class="card" style="padding:13px">
          <div style="font-size:10.5px;color:var(--muted);text-transform:uppercase;
               letter-spacing:.8px;font-weight:750">Different people</div>
          <div style="font-size:21px;font-weight:800;color:var(--accent)">${d.diff_avg ?? '—'}</div>
          <div style="font-size:11px;color:var(--faint)">closest ${d.diff_closest ?? '—'} · higher is better</div>
        </div>
      </div>

      ${d.margin !== null ? `<div class="card" style="padding:13px;margin-bottom:12px">
        <div style="display:flex;justify-content:space-between;font-size:12.5px">
          <span style="color:var(--muted)">Safety margin</span>
          <b style="color:${d.margin >= 0.08 ? 'var(--ok)' : 'var(--warn)'}">${d.margin}</b>
        </div>
        <div style="font-size:11px;color:var(--faint);margin-top:5px;line-height:1.5">
          The gap between the worst same-person match and the closest
          different-person match. Above 0.08 is comfortable.
        </div></div>` : ''}

      ${d.risky.length ? `<div class="sec" style="margin-top:0">Could be confused</div>
        ${d.risky.map(r => `<div style="display:flex;justify-content:space-between;
           padding:8px 0;border-bottom:1px solid var(--stroke);font-size:12.5px">
           <span>${esc(r.a)} ↔ ${esc(r.b)}</span>
           <b style="color:var(--bad)">${r.distance}</b></div>`).join('')}
        <p style="font-size:11.5px;color:var(--muted);margin-top:10px;line-height:1.55">
          Add more photos of these students, or lower the face threshold in Settings.
        </p>` : ''}

      ${d.thin.length ? `<div class="sec">Needs more photos</div>
        <p style="font-size:12.5px;color:var(--muted);line-height:1.6">
          ${d.thin.map(t => `${esc(t.name)} (${t.photos})`).join(', ')}
          <br><span style="color:var(--faint);font-size:11.5px">Three or more
          photos per student, from different angles, makes a real difference.</span>
        </p>` : ''}

      <div class="acts"><button class="btn" onclick="closeModal()">Close</button></div>`);
  } catch (e) { toast('Test failed', e.message, 'bad'); }
}


/* ══════════════════════════════════════════════════════════════════
   SCHOOL TIMETABLE
   ══════════════════════════════════════════════════════════════════ */
const DAY_NAMES = { mon:'Monday', tue:'Tuesday', wed:'Wednesday',
                    thu:'Thursday', fri:'Friday', sat:'Saturday', sun:'Sunday' };
const STATUS_COLOUR = {
  present:'var(--ok)', late:'var(--warn)', half_day:'var(--warn)',
  absent:'var(--bad)', too_early:'var(--accent3)',
  holiday:'var(--accent2)', closed:'var(--faint)',
};

function paintWindow(a) {
  const box = document.getElementById('windowNow');
  if (!box || !a) return;
  if (!a.enforced) { box.style.display = 'none'; return; }
  box.style.display = '';
  document.getElementById('windowDot').style.background =
    STATUS_COLOUR[a.status] || 'var(--faint)';
  document.getElementById('windowLabel').textContent = a.label;
  document.getElementById('windowSub').textContent = a.can_mark
    ? `Scans now count as ${a.label.toLowerCase()} · ${a.window}`
    : `Attendance cannot be marked right now · ${a.window}`;
}

async function openSchedule() {
  let d;
  try { d = await api('/schedule'); } catch (e) { return toast('Could not load', e.message, 'bad'); }
  const days = d.schedule.days || {};
  const holidays = (d.schedule.holidays || []).join(String.fromCharCode(10));

  modal(`
    <h2>🕘 School timetable</h2>
    <p class="sub">Four times per day decide the status. Anyone who has not
       scanned by closing time is recorded absent.</p>

    <div class="card" style="padding:12px;margin-bottom:12px;font-size:11.5px;
         color:var(--muted);line-height:1.7">
      <b style="color:var(--ok)">Open → Late</b> counts as on time ·
      <b style="color:var(--warn)">Late → Half</b> counts as late ·
      <b style="color:var(--warn)">Half → Close</b> counts as a half day ·
      <b style="color:var(--bad)">after Close</b> is absent
    </div>

    <div style="max-height:40vh;overflow-y:auto">
      <table style="width:100%;font-size:12px">
        <thead><tr>${['Day','On','Opens','Late','Half','Closes'].map(h =>
          `<th style="text-align:left;padding:5px;color:var(--muted);font-size:10px">${h}</th>`).join('')}</tr></thead>
        <tbody>${Object.keys(DAY_NAMES).map(k => {
          const c = days[k] || {};
          return `<tr data-day="${k}">
            <td style="padding:5px">${DAY_NAMES[k].slice(0,3)}</td>
            <td style="padding:5px"><input type="checkbox" data-f="active"
                ${c.active ? 'checked' : ''}></td>
            ${['start','late','half','close'].map(f =>
              `<td style="padding:4px"><input data-f="${f}" value="${esc(c[f]||'')}"
                 style="width:64px;padding:5px 6px;font-size:11.5px;
                        border-radius:8px;border:1px solid var(--stroke);
                        background:var(--surface);color:var(--text)"></td>`).join('')}
          </tr>`;
        }).join('')}</tbody>
      </table>
    </div>

    <div class="field" style="margin-top:14px">
      <label>Holidays — one date per line (YYYY-MM-DD)</label>
      <textarea id="schHolidays" rows="4"
        placeholder="2026-10-02&#10;2026-12-25">${esc(holidays)}</textarea>
    </div>

    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="schSave">Save timetable</button>
    </div>`);

  document.getElementById('schSave').onclick = async () => {
    const out = {};
    document.querySelectorAll('[data-day]').forEach(row => {
      const rec = {};
      row.querySelectorAll('[data-f]').forEach(el => {
        rec[el.dataset.f] = el.type === 'checkbox' ? el.checked : el.value.trim();
      });
      out[row.dataset.day] = rec;
    });
    const hol = document.getElementById('schHolidays').value
      .split(String.fromCharCode(10)).map(x => x.trim()).filter(Boolean);
    try {
      const r = await api('/schedule', { method:'POST',
        body:JSON.stringify({ days: out, holidays: hol }) });
      toast('Timetable saved', r.now.label, 'ok');
      paintWindow(r.now);
      closeModal(); refresh();
    } catch (e) { toast('Save failed', e.message, 'bad'); }
  };
}

async function closeRegister() {
  modal(`
    <h2>📕 Close the register</h2>
    <p class="sub">Everyone who has not scanned today will be recorded as
       absent, with a timestamp in the audit log. This cannot be undone,
       though you can still mark someone afterwards.</p>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn bad" id="crGo">Close register</button>
    </div>`);
  document.getElementById('crGo').onclick = async () => {
    try {
      const r = await api('/register/close', { method:'POST', body:'{}' });
      toast('Register closed',
            `${r.present} present · ${r.absent.length} absent`, 'ok');
      closeModal(); refresh();
    } catch (e) { toast('Failed', e.message, 'bad'); }
  };
}


/* ══════════════════════════════════════════════════════════════════
   STAFF ACCOUNTS, EXHIBITION MODE, BACKUPS
   ══════════════════════════════════════════════════════════════════ */
async function openStaff() {
  let d;
  try { d = await api('/users'); }
  catch (e) { return toast('Admins only', e.message, 'warn'); }

  const row = u => `
    <div class="card" style="padding:12px;margin-bottom:8px">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px">
        <b style="font-size:13.5px">${esc(u.display || u.username)}</b>
        <span style="color:var(--faint);font-size:11.5px">@${esc(u.username)}</span>
        <span class="pill ${u.role === 'admin' ? 'xp' : 'on'}"
              style="margin-left:auto">${u.role}</span>
      </div>
      <div class="grid2" style="gap:8px">
        <div class="field" style="margin:0"><label>Role</label>
          <select data-u="${esc(u.username)}" data-f="role">
            <option value="teacher" ${u.role==='teacher'?'selected':''}>Teacher</option>
            <option value="admin" ${u.role==='admin'?'selected':''}>Admin</option>
          </select></div>
        <div class="field" style="margin:0"><label>Classes (comma separated)</label>
          <input data-u="${esc(u.username)}" data-f="classes"
                 value="${esc((u.classes||[]).join(', '))}"
                 placeholder="${u.role === 'admin' ? 'all classes' : '8C, 9A'}"></div>
      </div>
      <div class="row" style="margin-top:8px">
        <button class="btn sm" data-save="${esc(u.username)}">Save</button>
        <button class="btn sm" data-pw="${esc(u.username)}">Reset password</button>
        <button class="btn sm bad" data-del="${esc(u.username)}">Remove</button>
      </div>
    </div>`;

  modal(`
    <h2>👥 Staff accounts</h2>
    <p class="sub">Teachers only see and mark the classes listed against them.
       Leave classes empty to give access to everyone.</p>
    <div style="max-height:44vh;overflow-y:auto">${d.users.map(row).join('')}</div>

    <div class="sec">Add a teacher</div>
    <div class="grid2">
      <div class="field"><label>Name</label><input id="nuDisplay" placeholder="Ms Sharma"></div>
      <div class="field"><label>Username</label><input id="nuUser" placeholder="sharma"></div>
    </div>
    <div class="grid2">
      <div class="field"><label>Password</label><input id="nuPass" type="password"></div>
      <div class="field"><label>Classes</label><input id="nuClasses"
        placeholder="${esc((d.grades||[]).slice(0,2).join(', ')) || '8C, 9A'}"></div>
    </div>
    <button class="btn primary block" id="nuAdd">Create account</button>

    <div class="acts"><button class="btn" onclick="closeModal()">Close</button></div>`);

  const parseClasses = v => v.split(',').map(x => x.trim()).filter(Boolean);

  document.querySelectorAll('[data-save]').forEach(b => b.onclick = async () => {
    const un = b.dataset.save;
    const role = document.querySelector(`[data-u="${un}"][data-f="role"]`).value;
    const cls = parseClasses(document.querySelector(`[data-u="${un}"][data-f="classes"]`).value);
    try {
      await api('/users/' + encodeURIComponent(un), { method:'PUT',
        body:JSON.stringify({ role, classes: cls }) });
      toast('Saved', un, 'ok'); openStaff();
    } catch (e) { toast('Could not save', e.message, 'bad'); }
  });

  document.querySelectorAll('[data-pw]').forEach(b => b.onclick = async () => {
    const pw = prompt(`New password for ${b.dataset.pw} (at least 6 characters)`);
    if (!pw) return;
    try {
      await api('/users/' + encodeURIComponent(b.dataset.pw), { method:'PUT',
        body:JSON.stringify({ password: pw }) });
      toast('Password reset', b.dataset.pw, 'ok');
    } catch (e) { toast('Failed', e.message, 'bad'); }
  });

  document.querySelectorAll('[data-del]').forEach(b => b.onclick = async () => {
    if (!confirm(`Remove the account "${b.dataset.del}"?`)) return;
    try {
      await api('/users/' + encodeURIComponent(b.dataset.del), { method:'DELETE' });
      toast('Removed', b.dataset.del, 'ok'); openStaff();
    } catch (e) { toast('Could not remove', e.message, 'bad'); }
  });

  document.getElementById('nuAdd').onclick = async () => {
    try {
      await api('/users', { method:'POST', body:JSON.stringify({
        username: document.getElementById('nuUser').value.trim(),
        password: document.getElementById('nuPass').value,
        display: document.getElementById('nuDisplay').value.trim(),
        role: 'teacher',
        classes: parseClasses(document.getElementById('nuClasses').value),
      })});
      toast('Teacher added', '', 'ok'); openStaff();
    } catch (e) { toast('Could not add', e.message, 'bad'); }
  };
}

async function refreshDemo() {
  try {
    const d = await api('/demo');
    const banner = document.getElementById('demoBanner');
    if (banner) {
      banner.style.display = d.active ? '' : 'none';
      if (d.active && d.info.started) {
        document.getElementById('demoInfo').textContent =
          `Snapshot taken ${d.info.started.replace('T',' ')} · ` +
          `${d.info.students} student(s). Restore when the demo is over.`;
      }
    }
    const b = d.backups || {};
    if (document.getElementById('bkLast'))
      document.getElementById('bkLast').textContent =
        b.last_date ? `${b.last_date} — ${b.last_result}` : 'not yet';
    if (document.getElementById('bkKept'))
      document.getElementById('bkKept').textContent =
        `${b.kept ?? 0} of ${b.keep_limit ?? 14}`;
  } catch (e) {}
}

function initDay() {
  const on = (id, fn) => { const el = document.getElementById(id); if (el) el.onclick = fn; };

  on('btnSlip', makeSlip);
  on('btnAlerts', openAlerts);
  on('btnBulkPhotos', bulkPhotos);
  on('btnPortal', openPortalInfo);
  on('btnStaff', openStaff);

  on('btnDemoStart', async () => {
    modal(`<h2>🎪 Start demo mode</h2>
      <p class="sub">Takes a snapshot of the roster, faces, scans and badges.
         Visitors can then enrol and scan without touching your real data.</p>
      <div class="acts">
        <button class="btn" onclick="closeModal()">Cancel</button>
        <button class="btn primary" id="dmGo">Take snapshot</button>
      </div>`);
    document.getElementById('dmGo').onclick = async () => {
      try {
        const r = await api('/demo/start', { method:'POST', body:'{}' });
        toast('Demo mode on', `${r.students} student(s) snapshotted`, 'ok');
        closeModal(); refreshDemo();
      } catch (e) { toast('Failed', e.message, 'bad'); }
    };
  });

  on('btnDemoRestore', async () => {
    modal(`<h2>↩ Restore the real roster</h2>
      <p class="sub">Everything added or scanned since the snapshot is
         discarded, and the roster returns to exactly how it was.</p>
      <div class="acts">
        <button class="btn" onclick="closeModal()">Cancel</button>
        <button class="btn ok" id="drGo">Restore</button>
      </div>`);
    document.getElementById('drGo').onclick = async () => {
      try {
        const r = await api('/demo/restore', { method:'POST', body:'{}' });
        toast('Restored', `${r.students} student(s)`, 'ok');
        closeModal(); refreshDemo(); refresh();
      } catch (e) { toast('Failed', e.message, 'bad'); }
    };
  });

  on('btnClearToday', async () => {
    try {
      const r = await api('/demo/clear-today', { method:'POST', body:'{}' });
      toast("Today's scans cleared", `${r.cleared} removed`, 'ok');
      refresh();
    } catch (e) { toast('Failed', e.message, 'bad'); }
  });

  on('btnBackupNow', async () => {
    try {
      await api('/backup/auto', { method:'POST', body:'{}' });
      toast('Backup saved', '', 'ok'); refreshDemo();
    } catch (e) { toast('Failed', e.message, 'bad'); }
  });
}

/* Hide admin-only controls from a teacher, so they are not offered
   buttons that will simply refuse them. */
function applyRole(d) {
  const admin = d.role === 'admin';
  ['btnStaff','btnDemoStart','btnDemoRestore','btnBackupNow','swAutoBackup',
   'btnBulkPhotos',
   'btnSchedule','swEnforce','btnCameras','btnCloseRegister','btnEncrypt',
   'btnRemove','btnImport','btnSection'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = admin ? '' : 'none';
  });
  const sec = document.getElementById('staffSec');
  if (sec) sec.style.display = admin ? '' : 'none';
  const who = document.getElementById('whoami');
  if (who && d.role) {
    const cls = (d.my_classes || []).join(', ');
    who.textContent = `${currentUser?.display || ''} (${d.role}${cls ? ' · ' + cls : ''})`;
  }
}


/* ══════════════════════════════════════════════════════════════════
   SLIPS · ALERTS · BULK PHOTOS · PARENT PORTAL
   ══════════════════════════════════════════════════════════════════ */
async function makeSlip() {
  const sel = selectedStudents();
  if (sel.length !== 1)
    return toast('Select one student', 'Click their row first', 'warn');
  const st = sel[0];
  modal(`
    <h2>🧾 Late arrival slip</h2>
    <p class="sub">A printable slip for ${esc(st.Name)} to hand to their teacher.</p>
    <div class="field"><label>Reason (optional)</label>
      <input id="slipReason" placeholder="Bus was late"></div>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="slipGo">Create slip</button>
    </div>`);
  document.getElementById('slipGo').onclick = async () => {
    try {
      const r = await api('/slip', { method:'POST', body:JSON.stringify({
        uid: String(st.UID),
        reason: document.getElementById('slipReason').value.trim() })});
      modal(`<h2>🧾 Slip ready</h2>
        <p class="sub">${esc(st.Name)} — saved to ${esc(r.folder)}</p>
        <div class="acts">
          <button class="btn" onclick="closeModal()">Close</button>
          <a class="btn primary" target="_blank" style="text-decoration:none"
             href="/api/slip/${encodeURIComponent(r.file)}">Open &amp; print ↗</a>
        </div>`);
    } catch (e) { toast('Could not create slip', e.message, 'bad'); }
  };
}

async function openAlerts() {
  const pct = parseFloat(document.getElementById('alertPct')?.value) || 85;
  try {
    const d = await api(`/alerts?pct=${pct}`);
    modal(`
      <h2>⚠ Attendance concerns</h2>
      <p class="sub">${d.students.length} of ${d.total} students are below
         ${d.threshold}% across ${d.school_days} school day(s).</p>
      ${d.students.length ? `<div style="max-height:48vh;overflow-y:auto">
        <table style="width:100%;font-size:12.5px">
          <thead><tr>${['Student','Class','Days','%'].map(h =>
            `<th style="text-align:left;padding:7px;color:var(--muted);font-size:10.5px">${h}</th>`).join('')}</tr></thead>
          <tbody>${d.students.map(r => `<tr style="border-bottom:1px solid var(--stroke)">
            <td style="padding:8px">${esc(r.name)}</td>
            <td style="padding:8px;color:var(--muted)">${esc(r.grade||'—')}</td>
            <td style="padding:8px">${r.days_present}/${d.school_days}</td>
            <td style="padding:8px"><span class="pill ${r.pct >= 75 ? 'late' : 'off'}">${r.pct}%</span></td>
          </tr>`).join('')}</tbody></table></div>`
        : '<p style="color:var(--ok);font-size:13.5px">Everyone is above the threshold.</p>'}
      <div class="acts">
        <button class="btn" onclick="closeModal()">Close</button>
        ${d.students.length ? '<button class="btn warn" id="alertMail">Email their parents</button>' : ''}
      </div>`);
    const mail = document.getElementById('alertMail');
    if (mail) mail.onclick = async () => {
      if (!confirm(`Email the parents of ${d.students.length} student(s)?`)) return;
      try {
        const r = await api('/alerts/notify', { method:'POST',
          body:JSON.stringify({ pct }) });
        toast('Emails sent', `${r.sent.length} sent, ${r.skipped.length} skipped`, 'ok');
        closeModal();
      } catch (e) { toast('Send failed', e.message, 'bad'); }
    };
  } catch (e) { toast('Could not load', e.message, 'bad'); }
}

function bulkPhotos() {
  modal(`
    <h2>🖼 Import photos in bulk</h2>
    <p class="sub">Pick a folder of face photos. Each file should be named
       after the student's roll number or UID — <b>8C01.jpg</b>, <b>F003.png</b>.
       Extra photos of a student already enrolled are added as further angles,
       which improves recognition.</p>
    <div class="field">
      <label>Photos</label>
      <input type="file" id="bpFiles" accept="image/*" multiple>
    </div>
    <div id="bpOut" style="font-size:12.5px;color:var(--muted);line-height:1.6"></div>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn primary" id="bpGo">Import</button>
    </div>`);
  document.getElementById('bpGo').onclick = async () => {
    const files = document.getElementById('bpFiles').files;
    if (!files.length) return toast('Pick some photos first', '', 'warn');
    const fd = new FormData();
    for (const f of files) fd.append('photos', f);
    document.getElementById('bpOut').textContent =
      `Encoding ${files.length} photo(s)… this can take a minute.`;
    try {
      const tok = localStorage.getItem('fa.token');
      const r = await fetch('/api/photos/bulk', { method:'POST', body:fd,
        headers: tok ? { 'Authorization': 'Bearer ' + tok } : {} });
      const d = await r.json();
      if (!r.ok) throw new Error(d.error);
      document.getElementById('bpOut').innerHTML =
        `<b style="color:var(--ok)">${d.added} photo(s) enrolled.</b>` +
        (d.no_match.length ? `<br><span style="color:var(--warn)">No student matched:
          ${d.no_match.map(esc).join(', ')}</span>` : '') +
        (d.no_face.length ? `<br><span style="color:var(--bad)">No face found in:
          ${d.no_face.map(esc).join(', ')}</span>` : '');
      toast('Bulk import done', `${d.added} enrolled`, 'ok');
      refresh();
    } catch (e) {
      document.getElementById('bpOut').textContent = '';
      toast('Import failed', e.message, 'bad');
    }
  };
}

function openPortalInfo() {
  const url = location.origin + '/portal';
  modal(`
    <h2>🔗 Parent portal</h2>
    <p class="sub">A read-only page showing one child's attendance, badges and
       recorded absences. Parents enter the code from their child's QR pass.</p>
    <div class="card" style="padding:14px;margin-bottom:12px">
      <div style="font-size:11px;color:var(--muted);text-transform:uppercase;
           letter-spacing:.8px;font-weight:750;margin-bottom:6px">Address</div>
      <div style="font-family:var(--mono);font-size:13px;word-break:break-all">${esc(url)}</div>
    </div>
    <p style="font-size:11.5px;color:var(--faint);line-height:1.6">
      This works on any device on the same network. It shows only the child
      whose code is entered — a guessed code cannot be used to browse the school.
    </p>
    <div class="acts">
      <button class="btn" onclick="closeModal()">Close</button>
      <a class="btn primary" href="/portal" target="_blank"
         style="text-decoration:none">Open portal ↗</a>
    </div>`);
}


/* ══════════════════════════════════════════════════════════════════
   DEVELOPER TOOLS
   The arcade is not part of the product. It stays completely invisible
   — no menu entry, no command, no hint in the help text — and only
   appears for whoever knows the unlock sequence.
   ══════════════════════════════════════════════════════════════════ */
let devUnlocked = false;
const DEV_WORD = 'devmode';
let devBuffer = '';

function unlockDev(silent) {
  devUnlocked = true;
  sessionStorage.setItem('fa.dev', '1');
  if (!silent) {
    toast('Developer tools unlocked', 'Ctrl+Shift+G opens the arcade', 'xp');
    if (typeof beep === 'function') beep([523, 659, 784], .09);
  }
}

function lockDev() {
  devUnlocked = false;
  sessionStorage.removeItem('fa.dev');
  closeArcade();
  toast('Developer tools locked', '', 'info');
}

function initDev() {
  if (sessionStorage.getItem('fa.dev') === '1') devUnlocked = true;

  document.addEventListener('keydown', e => {
    // Ctrl+Shift+G, once unlocked
    if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'g') {
      e.preventDefault();
      if (devUnlocked) openArcade();
      return;
    }
    if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'l' && devUnlocked) {
      e.preventDefault(); lockDev(); return;
    }
    // Typing the unlock word anywhere outside a text field
    if (e.target.matches('input,textarea,[contenteditable]')) return;
    if (e.key.length !== 1) return;
    devBuffer = (devBuffer + e.key.toLowerCase()).slice(-DEV_WORD.length);
    if (devBuffer === DEV_WORD && !devUnlocked) { devBuffer = ''; unlockDev(); }
  });
}

window.unlockDev = unlockDev;
window.lockDev = lockDev;
window.openStaff = openStaff;
window.switchAccount = switchAccount;
window.signOut = signOut;
window.closeModal = closeModal;
window.showQr = showQr;
window.openCalendar = openCalendar;
document.addEventListener('DOMContentLoaded', init);


/* ══════════════════════════════════════════════════════════════════
   FaceAttend Arcade — five mini games
   Opened by telling the AI Buddy you're bored.
   ══════════════════════════════════════════════════════════════════ */
'use strict';

const GAMES = [
  { id:'snake', name:'Snake',        ic:'🐍', desc:'Eat, grow, do not bite yourself.' },
  { id:'trex',  name:'Dino Run',     ic:'🦖', desc:'The offline runner. Jump the cacti.' },
  { id:'pac',   name:'Pac-Attend',   ic:'👻', desc:'Collect every pellet, dodge the ghosts.' },
  { id:'chess', name:'Chess',        ic:'♟',  desc:'Full rules, thinking opponent.' },
  { id:'liar',  name:'Spot the Liar',ic:'🕵',  desc:'One student is lying. Work out who.' },
  { id:'dungeon',name:'Dungeon Crawler',ic:'⚔', desc:'Survive the waves. Melee and shoot.' },
  { id:'pong',  name:'Pong',          ic:'🏓', desc:'The original. First to eleven.' },
  { id:'ttt',   name:'Tic-Tac-Toe',   ic:'⭕', desc:'Three in a row, unbeatable opponent.' },
  { id:'memory',name:'Memory Match',  ic:'🧠', desc:'Flip the pairs, beat your best time.' },
  { id:'2048',  name:'2048',          ic:'🔢', desc:'Slide, merge, reach the 2048 tile.' },
];

let arcadeCleanup = null;

function openArcade(gameId) {
  if (!devUnlocked) return;          // not part of the product
  closeGame();
  const veil = document.getElementById('arcadeVeil');
  veil.classList.add('on');
  gameId ? launchGame(gameId) : showArcadeMenu();
}

function closeArcade() {
  closeGame();
  document.getElementById('arcadeVeil').classList.remove('on');
}

function closeGame() {
  if (arcadeCleanup) { try { arcadeCleanup(); } catch (e) {} arcadeCleanup = null; }
  // Cleared here as well as on the overlay, so the guard holds even if a
  // game is mounted directly rather than through openArcade().
  window.__gameActive = false;
}

function showArcadeMenu() {
  closeGame();          // also clears __gameActive
  document.getElementById('arcadeBox').innerHTML = `
    <div class="arcade-head">
      <div>
        <h2>🕹 Arcade</h2>
        <p>Ten games. Pick one — your attendance data is safe, I promise.</p>
      </div>
      <button class="btn sm" onclick="closeArcade()">✕ Close</button>
    </div>
    <div class="gamegrid">
      ${GAMES.map(g => `
        <div class="gamecard" data-g="${g.id}">
          <div class="gi">${g.ic}</div>
          <b>${g.name}</b>
          <span>${g.desc}</span>
          <div class="hs" id="hs-${g.id}"></div>
        </div>`).join('')}
    </div>`;
  document.querySelectorAll('.gamecard').forEach(c =>
    c.onclick = () => launchGame(c.dataset.g));
  GAMES.forEach(g => {
    const best = localStorage.getItem('fa.hs.' + g.id);
    if (best) document.getElementById('hs-' + g.id).textContent = 'Best: ' + best;
  });
}

function launchGame(id) {
  closeGame();
  window.__gameActive = true;
  const g = GAMES.find(x => x.id === id);
  document.getElementById('arcadeBox').innerHTML = `
    <div class="arcade-head">
      <div><h2>${g.ic} ${g.name}</h2><p id="gameStatus">${g.desc}</p></div>
      <div style="display:flex;gap:8px">
        <button class="btn sm" onclick="showArcadeMenu()">← Games</button>
        <button class="btn sm" onclick="closeArcade()">✕ Close</button>
      </div>
    </div>
    <div id="gameStage" class="gamestage"></div>`;
  ({ snake:gSnake, trex:gTrex, pac:gPac, chess:gChess, liar:gLiar,
     dungeon:gDungeon, pong:gPong, ttt:gTtt, memory:gMemory,
     '2048':g2048 })[id]();
}

const stage = () => document.getElementById('gameStage');
const setStatus = t => { const el = document.getElementById('gameStatus'); if (el) el.textContent = t; };

function saveHigh(id, score) {
  const k = 'fa.hs.' + id;
  const best = +(localStorage.getItem(k) || 0);
  if (score > best) { localStorage.setItem(k, score); return true; }
  return false;
}

/* ════════════════════════════════════════════════════════════════
   1. SNAKE
   ════════════════════════════════════════════════════════════════ */
function gSnake() {
  const N = 21, CELL = 22, W = N * CELL;
  stage().innerHTML = `<canvas id="cv" width="${W}" height="${W}"></canvas>
    <div class="ghint">Arrow keys or WASD · <kbd>Space</kbd> pause</div>`;
  const cv = document.getElementById('cv'), cx = cv.getContext('2d');

  let snake = [{ x:10, y:10 }], dir = { x:1, y:0 }, next = { x:1, y:0 };
  let food = rndCell(), score = 0, dead = false, paused = false, tick = 0;

  function rndCell() {
    let p;
    do { p = { x:(Math.random()*N)|0, y:(Math.random()*N)|0 }; }
    while (snake.some(s => s.x === p.x && s.y === p.y));
    return p;
  }

  const key = e => {
    const k = e.key.toLowerCase();
    if (k === ' ') { paused = !paused; setStatus(paused ? 'Paused' : 'Score ' + score); e.preventDefault(); return; }
    const m = { arrowup:[0,-1], w:[0,-1], arrowdown:[0,1], s:[0,1],
                arrowleft:[-1,0], a:[-1,0], arrowright:[1,0], d:[1,0] }[k];
    if (!m) return;
    e.preventDefault();
    if (m[0] === -dir.x && m[1] === -dir.y) return;   // no instant reverse
    next = { x:m[0], y:m[1] };
  };
  window.addEventListener('keydown', key);

  function step() {
    if (dead || paused) return;
    dir = next;
    const head = { x:snake[0].x + dir.x, y:snake[0].y + dir.y };
    if (head.x < 0 || head.y < 0 || head.x >= N || head.y >= N ||
        snake.some(s => s.x === head.x && s.y === head.y)) {
      dead = true;
      const rec = saveHigh('snake', score);
      setStatus(`Game over — ${score} point${score === 1 ? '' : 's'}${rec ? ' · new best!' : ''}. Press R to retry.`);
      if (rec && typeof confetti === 'function') confetti(40);
      return;
    }
    snake.unshift(head);
    if (head.x === food.x && head.y === food.y) {
      score++; food = rndCell(); setStatus('Score ' + score);
      if (typeof beep === 'function') beep([740, 940], .07);
    } else snake.pop();
  }

  function draw() {
    cx.fillStyle = '#080b12'; cx.fillRect(0, 0, W, W);
    cx.strokeStyle = 'rgba(255,255,255,.035)';
    for (let i = 1; i < N; i++) {
      cx.beginPath(); cx.moveTo(i*CELL,0); cx.lineTo(i*CELL,W); cx.stroke();
      cx.beginPath(); cx.moveTo(0,i*CELL); cx.lineTo(W,i*CELL); cx.stroke();
    }
    // food with a soft pulse
    const p = 3 + Math.sin(tick/7) * 1.6;
    cx.fillStyle = '#ff5d73';
    cx.beginPath(); cx.arc(food.x*CELL+CELL/2, food.y*CELL+CELL/2, CELL/2-p, 0, 7); cx.fill();
    snake.forEach((s, i) => {
      const t = i / Math.max(1, snake.length);
      cx.fillStyle = i === 0 ? '#5b8cff'
        : `rgb(${Math.round(91-t*40)},${Math.round(140-t*50)},${Math.round(255-t*90)})`;
      rr(cx, s.x*CELL+2, s.y*CELL+2, CELL-4, CELL-4, 6); cx.fill();
    });
    if (dead) {
      cx.fillStyle = 'rgba(5,7,13,.8)'; cx.fillRect(0, 0, W, W);
      cx.fillStyle = '#eef2fb'; cx.font = '700 26px system-ui'; cx.textAlign = 'center';
      cx.fillText('Game over', W/2, W/2 - 6);
      cx.font = '400 14px system-ui'; cx.fillStyle = '#94a3bb';
      cx.fillText('Press R to play again', W/2, W/2 + 22);
    }
  }

  const retry = e => {
    if (e.key.toLowerCase() === 'r' && dead) {
      snake = [{x:10,y:10}]; dir = {x:1,y:0}; next = {x:1,y:0};
      food = rndCell(); score = 0; dead = false; setStatus('Score 0');
    }
  };
  window.addEventListener('keydown', retry);

  let acc = 0, last = performance.now(), raf;
  (function loop(now) {
    raf = requestAnimationFrame(loop);
    acc += now - last; last = now; tick++;
    const speed = Math.max(70, 135 - score * 3);
    if (acc > speed) { acc = 0; step(); }
    draw();
  })(last);

  setStatus('Score 0');
  arcadeCleanup = () => {
    cancelAnimationFrame(raf);
    window.removeEventListener('keydown', key);
    window.removeEventListener('keydown', retry);
  };
}

function rr(cx, x, y, w, h, r) {
  cx.beginPath();
  cx.moveTo(x+r, y); cx.arcTo(x+w, y, x+w, y+h, r);
  cx.arcTo(x+w, y+h, x, y+h, r); cx.arcTo(x, y+h, x, y, r);
  cx.arcTo(x, y, x+w, y, r); cx.closePath();
}

/* ════════════════════════════════════════════════════════════════
   2. T-REX RUNNER
   ════════════════════════════════════════════════════════════════ */
function gTrex() {
  const W = 640, H = 220, GROUND = H - 34;
  stage().innerHTML = `<canvas id="cv" width="${W}" height="${H}"></canvas>
    <div class="ghint"><kbd>Space</kbd> jump · <kbd>↓</kbd> duck</div>`;
  const cv = document.getElementById('cv'), cx = cv.getContext('2d');

  let y = 0, vy = 0, ducking = false, obstacles = [], clouds = [];
  let score = 0, speed = 6.2, dead = false, frame = 0, night = false;

  for (let i = 0; i < 4; i++) clouds.push({ x:Math.random()*W, y:22+Math.random()*60, s:.35+Math.random()*.4 });

  const key = e => {
    const k = e.key.toLowerCase();
    if (k === ' ' || k === 'arrowup' || k === 'w') {
      e.preventDefault();
      if (dead) return reset();
      if (y === 0) { vy = 12.6; if (typeof beep === 'function') beep([560], .05); }
    }
    if (k === 'arrowdown' || k === 's') { ducking = true; e.preventDefault(); }
  };
  const keyUp = e => {
    const k = e.key.toLowerCase();
    if (k === 'arrowdown' || k === 's') ducking = false;
  };
  window.addEventListener('keydown', key);
  window.addEventListener('keyup', keyUp);

  function reset() {
    y = 0; vy = 0; obstacles = []; score = 0; speed = 6.2; dead = false; night = false;
    setStatus('Score 0');
  }

  function spawn() {
    const flying = score > 260 && Math.random() < .28;
    obstacles.push(flying
      ? { x:W + 20, y:GROUND - 52 - (Math.random() < .5 ? 22 : 0), w:30, h:20, fly:true }
      : { x:W + 20, y:GROUND - 30, w:14 + (Math.random()*16|0), h:30, fly:false });
  }

  let raf, spawnAt = 60;
  (function loop() {
    raf = requestAnimationFrame(loop);
    frame++;
    if (!dead) {
      score++;
      if (score % 520 === 0) night = !night;
      speed = 6.2 + score / 420;
      vy -= 0.62; y += vy;
      if (y < 0) { y = 0; vy = 0; }
      if (--spawnAt <= 0) { spawn(); spawnAt = 48 + Math.random() * 62; }
      obstacles.forEach(o => o.x -= speed);
      obstacles = obstacles.filter(o => o.x > -60);

      const dx = 62, dw = ducking && y === 0 ? 46 : 34;
      const dh = ducking && y === 0 ? 24 : 42;
      const dy = GROUND - dh - y;
      for (const o of obstacles) {
        if (dx + dw - 6 > o.x && dx + 6 < o.x + o.w &&
            dy + dh - 4 > o.y && dy + 6 < o.y + o.h) {
          dead = true;
          const rec = saveHigh('trex', score);
          setStatus(`Crashed at ${score}${rec ? ' · new best!' : ''}. Space to retry.`);
          if (rec && typeof confetti === 'function') confetti(40);
        }
      }
    }

    // draw
    cx.fillStyle = night ? '#05070d' : '#0d1220'; cx.fillRect(0, 0, W, H);
    cx.fillStyle = night ? 'rgba(255,255,255,.10)' : 'rgba(255,255,255,.07)';
    clouds.forEach(c => {
      c.x -= c.s; if (c.x < -60) { c.x = W + 30; c.y = 22 + Math.random()*60; }
      cx.beginPath(); cx.ellipse(c.x, c.y, 26, 9, 0, 0, 7); cx.fill();
    });
    cx.strokeStyle = night ? '#2a3348' : '#37415a'; cx.lineWidth = 2;
    cx.beginPath(); cx.moveTo(0, GROUND); cx.lineTo(W, GROUND); cx.stroke();

    obstacles.forEach(o => {
      cx.fillStyle = o.fly ? '#b57bff' : '#2ee6a8';
      if (o.fly) {
        const flap = Math.sin(frame / 5) * 6;
        cx.beginPath();
        cx.moveTo(o.x, o.y + 10); cx.lineTo(o.x + 15, o.y + flap);
        cx.lineTo(o.x + 30, o.y + 10); cx.lineTo(o.x + 15, o.y + 16);
        cx.closePath(); cx.fill();
      } else {
        rr(cx, o.x, o.y, o.w, o.h, 3); cx.fill();
        rr(cx, o.x - 5, o.y + 8, 6, 12, 3); cx.fill();
        rr(cx, o.x + o.w - 1, o.y + 6, 6, 12, 3); cx.fill();
      }
    });

    const dw = ducking && y === 0 ? 46 : 34, dh = ducking && y === 0 ? 24 : 42;
    const dy = GROUND - dh - y;
    cx.fillStyle = dead ? '#ff5d73' : '#5b8cff';
    rr(cx, 62, dy, dw, dh, 6); cx.fill();
    cx.fillStyle = '#eef2fb';
    cx.fillRect(62 + dw - 12, dy + 7, 5, 5);
    if (y === 0 && !dead && (frame % 12 < 6)) {
      cx.fillStyle = dead ? '#ff5d73' : '#5b8cff';
      cx.fillRect(66, GROUND - 4, 8, 5);
    }

    cx.fillStyle = '#94a3bb'; cx.font = '600 14px ui-monospace,monospace';
    cx.textAlign = 'right'; cx.fillText(String(score).padStart(5, '0'), W - 14, 26);
    const best = localStorage.getItem('fa.hs.trex');
    if (best) { cx.fillStyle = '#56627a'; cx.fillText('HI ' + String(best).padStart(5,'0'), W - 90, 26); }

    if (dead) {
      cx.fillStyle = 'rgba(5,7,13,.72)'; cx.fillRect(0, 0, W, H);
      cx.fillStyle = '#eef2fb'; cx.font = '700 24px system-ui'; cx.textAlign = 'center';
      cx.fillText('G A M E   O V E R', W/2, H/2);
      cx.font = '400 13px system-ui'; cx.fillStyle = '#94a3bb';
      cx.fillText('press space', W/2, H/2 + 26);
    }
  })();

  setStatus('Score 0');
  arcadeCleanup = () => {
    cancelAnimationFrame(raf);
    window.removeEventListener('keydown', key);
    window.removeEventListener('keyup', keyUp);
  };
}

/* ════════════════════════════════════════════════════════════════
   3. PAC-ATTEND
   ════════════════════════════════════════════════════════════════ */
function gPac() {
  const MAP = [
    "###################",
    "#........#........#",
    "#o##.###.#.###.##o#",
    "#.................#",
    "#.##.#.#####.#.##.#",
    "#....#...#...#....#",
    "####.### # ###.####",
    "   #.#   G   #.#   ",
    "####.# ##-## #.####",
    "    .  #GGG#  .    ",
    "####.# ##### #.####",
    "   #.#       #.#   ",
    "####.# ##### #.####",
    "#........#........#",
    "#.##.###.#.###.##.#",
    "#o.#.....P.....#.o#",
    "##.#.#.#####.#.#.##",
    "#....#...#...#....#",
    "#.######.#.######.#",
    "#.................#",
    "###################",
  ];
  const CELL = 22, COLS = MAP[0].length, ROWS = MAP.length;
  const HUD = 26;
  const W = COLS * CELL, H = ROWS * CELL + HUD;

  stage().innerHTML = `<canvas id="cv" width="${W}" height="${H}"></canvas>
    <div class="ghint">Arrow keys or WASD · big dots let you eat ghosts</div>`;
  const cv = document.getElementById('cv'), cx = cv.getContext('2d');

  let grid = MAP.map(r => r.split(''));
  let pac, ghosts, score = 0, lives = 3, pellets = 0, frightened = 0, over = false, frame = 0;

  function init() {
    grid = MAP.map(r => r.split(''));
    ghosts = []; pellets = 0;
    for (let y = 0; y < ROWS; y++) for (let x = 0; x < COLS; x++) {
      const c = grid[y][x];
      if (c === 'P') { pac = { x, y, dx:0, dy:0, nx:0, ny:0, mx:0 }; grid[y][x] = '.'; pellets++; }
      if (c === 'G') {
        ghosts.push({ x, y, dx:0, dy:-1, home:{x,y},
                      col:['#ff5d73','#ffa23d','#22d3ee','#b57bff'][ghosts.length % 4] });
        grid[y][x] = ' ';
      }
      if (c === '.' || c === 'o') pellets++;
    }
  }
  init();

  const wall = (x, y) => {
    if (y < 0 || y >= ROWS) return true;
    x = (x + COLS) % COLS;
    return grid[y][x] === '#' || grid[y][x] === '-';
  };

  const key = e => {
    const k = e.key.toLowerCase();
    const m = { arrowup:[0,-1], w:[0,-1], arrowdown:[0,1], s:[0,1],
                arrowleft:[-1,0], a:[-1,0], arrowright:[1,0], d:[1,0] }[k];
    if (!m) return;
    e.preventDefault();
    pac.nx = m[0]; pac.ny = m[1];
  };
  window.addEventListener('keydown', key);

  function movePac() {
    if (!wall(pac.x + pac.nx, pac.y + pac.ny)) { pac.dx = pac.nx; pac.dy = pac.ny; }
    if (wall(pac.x + pac.dx, pac.y + pac.dy)) return;
    pac.x = (pac.x + pac.dx + COLS) % COLS; pac.y += pac.dy;
    const c = grid[pac.y][pac.x];
    if (c === '.') { grid[pac.y][pac.x] = ' '; score += 10; pellets--; setStatus('Score ' + score); }
    if (c === 'o') {
      grid[pac.y][pac.x] = ' '; score += 50; pellets--; frightened = 44;
      setStatus('Score ' + score + ' — ghosts are scared!');
      if (typeof beep === 'function') beep([420, 620], .09);
    }
    if (pellets <= 0) {
      over = true;
      saveHigh('pac', score);
      setStatus(`Cleared the board! ${score} points.`);
      if (typeof confetti === 'function') confetti(90);
    }
  }

  function moveGhosts() {
    ghosts.forEach(g => {
      const opts = [[0,-1],[0,1],[-1,0],[1,0]].filter(([dx,dy]) =>
        !wall(g.x+dx, g.y+dy) && !(dx === -g.dx && dy === -g.dy));
      const choices = opts.length ? opts : [[-g.dx, -g.dy]];
      // chase when normal, flee when frightened; some randomness either way
      let best = choices[(Math.random()*choices.length)|0];
      if (Math.random() < .72) {
        let bd = frightened ? -1e9 : 1e9;
        choices.forEach(([dx,dy]) => {
          const d = Math.hypot(g.x+dx-pac.x, g.y+dy-pac.y);
          if (frightened ? d > bd : d < bd) { bd = d; best = [dx,dy]; }
        });
      }
      g.dx = best[0]; g.dy = best[1];
      g.x = (g.x + g.dx + COLS) % COLS; g.y += g.dy;
    });
  }

  function collide() {
    ghosts.forEach(g => {
      if (g.x === pac.x && g.y === pac.y) {
        if (frightened) {
          score += 200; g.x = g.home.x; g.y = g.home.y;
          if (typeof beep === 'function') beep([880, 1180], .08);
        } else {
          lives--;
          if (typeof beep === 'function') beep([300, 180], .16);
          if (lives <= 0) {
            over = true; saveHigh('pac', score);
            setStatus(`Caught! Final score ${score}. Press R to retry.`);
          } else {
            pac.x = 9; pac.y = 15; pac.dx = pac.dy = pac.nx = pac.ny = 0;
            ghosts.forEach(gg => { gg.x = gg.home.x; gg.y = gg.home.y; });
            setStatus(`Ouch — ${lives} live${lives === 1 ? '' : 's'} left`);
          }
        }
      }
    });
  }

  const retry = e => {
    if (e.key.toLowerCase() === 'r' && over) {
      init(); score = 0; lives = 3; over = false; frightened = 0; setStatus('Score 0');
    }
  };
  window.addEventListener('keydown', retry);

  let acc = 0, last = performance.now(), gacc = 0, raf;
  (function loop(now) {
    raf = requestAnimationFrame(loop);
    const dt = now - last; last = now; frame++;
    if (!over) {
      acc += dt; gacc += dt;
      if (acc > 138) { acc = 0; movePac(); collide(); }
      if (gacc > (frightened ? 260 : 178)) { gacc = 0; moveGhosts(); collide(); }
      if (frightened > 0) frightened--;
    }

    cx.fillStyle = '#05070d'; cx.fillRect(0, 0, W, H);
    for (let y = 0; y < ROWS; y++) for (let x = 0; x < COLS; x++) {
      const c = grid[y][x], px = x*CELL, py = y*CELL;
      if (c === '#') {
        cx.fillStyle = '#16307a';
        rr(cx, px+2, py+2, CELL-4, CELL-4, 5); cx.fill();
        cx.strokeStyle = '#3a6bd8'; cx.lineWidth = 1.4;
        rr(cx, px+2.5, py+2.5, CELL-5, CELL-5, 5); cx.stroke();
      } else if (c === '-') {
        cx.fillStyle = '#b57bff'; cx.fillRect(px+2, py+CELL/2-2, CELL-4, 4);
      } else if (c === '.') {
        cx.fillStyle = '#ffd166';
        cx.beginPath(); cx.arc(px+CELL/2, py+CELL/2, 2.4, 0, 7); cx.fill();
      } else if (c === 'o') {
        cx.fillStyle = '#ffd166';
        cx.beginPath(); cx.arc(px+CELL/2, py+CELL/2, 5.4 + Math.sin(frame/8)*1.1, 0, 7); cx.fill();
      }
    }

    // pac
    const mouth = Math.abs(Math.sin(frame/5)) * .32 + .04;
    const ang = pac.dx === 1 ? 0 : pac.dx === -1 ? Math.PI
              : pac.dy === 1 ? Math.PI/2 : pac.dy === -1 ? -Math.PI/2 : 0;
    cx.fillStyle = '#ffd93d';
    cx.beginPath();
    cx.moveTo(pac.x*CELL+CELL/2, pac.y*CELL+CELL/2);
    cx.arc(pac.x*CELL+CELL/2, pac.y*CELL+CELL/2, CELL/2-2,
           ang+mouth*Math.PI, ang-mouth*Math.PI+Math.PI*2);
    cx.closePath(); cx.fill();

    // ghosts
    ghosts.forEach(g => {
      const gx = g.x*CELL+CELL/2, gy = g.y*CELL+CELL/2, r = CELL/2-2;
      cx.fillStyle = frightened ? (frightened < 12 && frame % 8 < 4 ? '#eef2fb' : '#3a6bd8') : g.col;
      cx.beginPath();
      cx.arc(gx, gy-1, r, Math.PI, 0);
      cx.lineTo(gx+r, gy+r-2);
      for (let i = 0; i < 3; i++) cx.lineTo(gx+r-(i*2+1)*r/3, gy+r-2-(i%2?0:4));
      cx.lineTo(gx-r, gy+r-2); cx.closePath(); cx.fill();
      if (!frightened) {
        cx.fillStyle = '#fff';
        cx.beginPath(); cx.arc(gx-3.5, gy-2, 2.6, 0, 7); cx.arc(gx+3.5, gy-2, 2.6, 0, 7); cx.fill();
        cx.fillStyle = '#0b0e14';
        cx.beginPath(); cx.arc(gx-3.5+g.dx*1.2, gy-2+g.dy*1.2, 1.3, 0, 7);
        cx.arc(gx+3.5+g.dx*1.2, gy-2+g.dy*1.2, 1.3, 0, 7); cx.fill();
      }
    });

    cx.fillStyle = '#05070d'; cx.fillRect(0, H - HUD, W, HUD);
    cx.fillStyle = '#eef2fb'; cx.font = '600 13px system-ui'; cx.textAlign = 'left';
    cx.fillText('Score ' + score, 8, H - 8);
    cx.fillStyle = '#ff5d73'; cx.textAlign = 'right';
    cx.fillText('♥'.repeat(Math.max(0, lives)), W - 8, H - 8);

    if (over) {
      cx.fillStyle = 'rgba(5,7,13,.78)'; cx.fillRect(0, 0, W, H);
      cx.fillStyle = '#eef2fb'; cx.font = '700 26px system-ui'; cx.textAlign = 'center';
      cx.fillText(pellets <= 0 ? 'Board cleared!' : 'Game over', W/2, (H-HUD)/2);
      cx.font = '400 13px system-ui'; cx.fillStyle = '#94a3bb';
      cx.fillText('press R to play again', W/2, (H-HUD)/2 + 26);
    }
  })(last);

  setStatus('Score 0');
  arcadeCleanup = () => {
    cancelAnimationFrame(raf);
    window.removeEventListener('keydown', key);
    window.removeEventListener('keydown', retry);
  };
}


/* ════════════════════════════════════════════════════════════════
   4. CHESS  — full legal moves, minimax + alpha-beta opponent
   ════════════════════════════════════════════════════════════════ */
function gChess() {
  const GLYPH = { P:'♙',N:'♘',B:'♗',R:'♖',Q:'♕',K:'♔',
                  p:'♟',n:'♞',b:'♝',r:'♜',q:'♛',k:'♚' };
  const VAL = { p:100, n:320, b:330, r:500, q:900, k:20000 };

  // piece-square tables make the engine play sensibly, not just grab material
  const PST = {
    p:[0,0,0,0,0,0,0,0, 50,50,50,50,50,50,50,50, 10,10,20,30,30,20,10,10,
       5,5,10,25,25,10,5,5, 0,0,0,20,20,0,0,0, 5,-5,-10,0,0,-10,-5,5,
       5,10,10,-20,-20,10,10,5, 0,0,0,0,0,0,0,0],
    n:[-50,-40,-30,-30,-30,-30,-40,-50, -40,-20,0,0,0,0,-20,-40,
       -30,0,10,15,15,10,0,-30, -30,5,15,20,20,15,5,-30,
       -30,0,15,20,20,15,0,-30, -30,5,10,15,15,10,5,-30,
       -40,-20,0,5,5,0,-20,-40, -50,-40,-30,-30,-30,-30,-40,-50],
    b:[-20,-10,-10,-10,-10,-10,-10,-20, -10,0,0,0,0,0,0,-10,
       -10,0,5,10,10,5,0,-10, -10,5,5,10,10,5,5,-10,
       -10,0,10,10,10,10,0,-10, -10,10,10,10,10,10,10,-10,
       -10,5,0,0,0,0,5,-10, -20,-10,-10,-10,-10,-10,-10,-20],
    r:[0,0,0,0,0,0,0,0, 5,10,10,10,10,10,10,5, -5,0,0,0,0,0,0,-5,
       -5,0,0,0,0,0,0,-5, -5,0,0,0,0,0,0,-5, -5,0,0,0,0,0,0,-5,
       -5,0,0,0,0,0,0,-5, 0,0,0,5,5,0,0,0],
    q:[-20,-10,-10,-5,-5,-10,-10,-20, -10,0,0,0,0,0,0,-10,
       -10,0,5,5,5,5,0,-10, -5,0,5,5,5,5,0,-5, 0,0,5,5,5,5,0,-5,
       -10,5,5,5,5,5,0,-10, -10,0,5,0,0,0,0,-10, -20,-10,-10,-5,-5,-10,-10,-20],
    k:[-30,-40,-40,-50,-50,-40,-40,-30, -30,-40,-40,-50,-50,-40,-40,-30,
       -30,-40,-40,-50,-50,-40,-40,-30, -30,-40,-40,-50,-50,-40,-40,-30,
       -20,-30,-30,-40,-40,-30,-30,-20, -10,-20,-20,-20,-20,-20,-20,-10,
       20,20,0,0,0,0,20,20, 20,30,10,0,0,10,30,20],
  };

  let board, turn, sel, legal, history, thinking, gameOver, lastMove;
  // ── modes ──────────────────────────────────────────────────────
  // difficulty sets search depth + how much the engine blunders;
  // twoPlayer disables the engine entirely for hot-seat games.
  const LEVELS = {
    easy:   { depth: 1, jitter: 90, label: 'Easy'   },
    normal: { depth: 2, jitter: 12, label: 'Normal' },
    hard:   { depth: 3, jitter: 0,  label: 'Hard'   },
  };
  let level = localStorage.getItem('fa.chess.level') || 'normal';
  let twoPlayer = localStorage.getItem('fa.chess.2p') === '1';
  let playAs = localStorage.getItem('fa.chess.side') || 'w';
  let undoStack = [];

  function reset() {
    board = [
      'rnbqkbnr'.split(''), 'pppppppp'.split(''),
      [' ',' ',' ',' ',' ',' ',' ',' '], [' ',' ',' ',' ',' ',' ',' ',' '],
      [' ',' ',' ',' ',' ',' ',' ',' '], [' ',' ',' ',' ',' ',' ',' ',' '],
      'PPPPPPPP'.split(''), 'RNBQKBNR'.split(''),
    ];
    turn = 'w'; sel = null; legal = []; history = []; thinking = false;
    gameOver = false; lastMove = null; undoStack = [];
  }
  reset();

  const isW = p => p !== ' ' && p === p.toUpperCase();
  const isB = p => p !== ' ' && p === p.toLowerCase();
  const mine = (p, t) => t === 'w' ? isW(p) : isB(p);
  const inB = (r, c) => r >= 0 && r < 8 && c >= 0 && c < 8;

  function pseudo(bd, t) {
    const out = [];
    const add = (r, c, nr, nc) => {
      if (!inB(nr, nc)) return false;
      const tgt = bd[nr][nc];
      if (tgt !== ' ' && mine(tgt, t)) return false;
      out.push({ r, c, nr, nc });
      return tgt === ' ';
    };
    for (let r = 0; r < 8; r++) for (let c = 0; c < 8; c++) {
      const p = bd[r][c];
      if (p === ' ' || !mine(p, t)) continue;
      const k = p.toLowerCase();
      if (k === 'p') {
        const d = t === 'w' ? -1 : 1, start = t === 'w' ? 6 : 1;
        if (inB(r+d, c) && bd[r+d][c] === ' ') {
          out.push({ r, c, nr:r+d, nc:c });
          if (r === start && bd[r+2*d][c] === ' ') out.push({ r, c, nr:r+2*d, nc:c });
        }
        for (const dc of [-1, 1]) {
          const nr = r+d, nc = c+dc;
          if (inB(nr, nc) && bd[nr][nc] !== ' ' && !mine(bd[nr][nc], t))
            out.push({ r, c, nr, nc });
        }
      } else if (k === 'n') {
        for (const [dr,dc] of [[-2,-1],[-2,1],[-1,-2],[-1,2],[1,-2],[1,2],[2,-1],[2,1]])
          add(r, c, r+dr, c+dc);
      } else if (k === 'k') {
        for (const [dr,dc] of [[-1,-1],[-1,0],[-1,1],[0,-1],[0,1],[1,-1],[1,0],[1,1]])
          add(r, c, r+dr, c+dc);
      } else {
        const dirs = k === 'b' ? [[-1,-1],[-1,1],[1,-1],[1,1]]
                   : k === 'r' ? [[-1,0],[1,0],[0,-1],[0,1]]
                   : [[-1,-1],[-1,1],[1,-1],[1,1],[-1,0],[1,0],[0,-1],[0,1]];
        for (const [dr,dc] of dirs) {
          let nr = r+dr, nc = c+dc;
          while (add(r, c, nr, nc)) { nr += dr; nc += dc; }
        }
      }
    }
    return out;
  }

  const apply = (bd, m) => {
    const n = bd.map(r => r.slice());
    let p = n[m.r][m.c];
    if (p.toLowerCase() === 'p' && (m.nr === 0 || m.nr === 7))
      p = isW(p) ? 'Q' : 'q';                     // auto-queen
    n[m.nr][m.nc] = p; n[m.r][m.c] = ' ';
    return n;
  };

  function attacked(bd, t) {
    const kc = t === 'w' ? 'K' : 'k';
    let kr = -1, kk = -1;
    for (let r = 0; r < 8; r++) for (let c = 0; c < 8; c++)
      if (bd[r][c] === kc) { kr = r; kk = c; }
    if (kr < 0) return true;
    return pseudo(bd, t === 'w' ? 'b' : 'w').some(m => m.nr === kr && m.nc === kk);
  }

  const legalMoves = (bd, t) => pseudo(bd, t).filter(m => !attacked(apply(bd, m), t));

  function evaluate(bd) {
    let s = 0;
    for (let r = 0; r < 8; r++) for (let c = 0; c < 8; c++) {
      const p = bd[r][c];
      if (p === ' ') continue;
      const k = p.toLowerCase(), w = isW(p);
      const idx = w ? r*8+c : (7-r)*8+c;
      const v = VAL[k] + (PST[k] ? PST[k][idx] : 0);
      s += w ? v : -v;
    }
    return s;
  }

  function search(bd, depth, alpha, beta, t) {
    if (depth === 0) return evaluate(bd);
    const moves = legalMoves(bd, t);
    if (!moves.length) return attacked(bd, t) ? (t === 'w' ? -99999 : 99999) : 0;
    // captures first — makes alpha-beta prune far more
    moves.sort((a, b) => (bd[b.nr][b.nc] !== ' ') - (bd[a.nr][a.nc] !== ' '));
    if (t === 'w') {
      let best = -1e9;
      for (const m of moves) {
        best = Math.max(best, search(apply(bd, m), depth-1, alpha, beta, 'b'));
        alpha = Math.max(alpha, best);
        if (beta <= alpha) break;
      }
      return best;
    }
    let best = 1e9;
    for (const m of moves) {
      best = Math.min(best, search(apply(bd, m), depth-1, alpha, beta, 'w'));
      beta = Math.min(beta, best);
      if (beta <= alpha) break;
    }
    return best;
  }

  function aiMove() {
    const engineSide = playAs === 'w' ? 'b' : 'w';
    const moves = legalMoves(board, engineSide);
    if (!moves.length) return null;
    moves.sort((a, b) => (board[b.nr][b.nc] !== ' ') - (board[a.nr][a.nc] !== ' '));
    let best = null, bestScore = playAs === 'w' ? 1e9 : -1e9;
    for (const m of moves) {
      const L = LEVELS[level] || LEVELS.normal;
      const sc = search(apply(board, m), L.depth, -1e9, 1e9,
                        playAs === 'w' ? 'w' : 'b')
               + (Math.random() * L.jitter - L.jitter / 2);
      const better = playAs === 'w' ? (sc < bestScore) : (sc > bestScore);
      if (best === null || better) { bestScore = sc; best = m; }
    }
    return best;
  }


  function undoMove() {
    if (!undoStack.length) return setStatus('Nothing to undo.');
    const snap = undoStack.pop();
    board = snap.board.map(r => r.slice());
    turn = snap.turn; lastMove = snap.lastMove;
    sel = null; legal = []; gameOver = false;
    render(); setStatus('Move undone.');
  }

  function showHint() {
    if (thinking || gameOver) return;
    const side = twoPlayer ? turn : playAs;
    if (turn !== side) return setStatus('Not your turn.');
    const moves = legalMoves(board, side);
    if (!moves.length) return setStatus('No legal moves.');
    // Score each move with the engine, from this side's point of view.
    let best = moves[0], bestScore = side === 'w' ? -1e9 : 1e9;
    for (const m of moves) {
      const sc = search(apply(board, m), 1, -1e9, 1e9, side === 'w' ? 'b' : 'w');
      if (side === 'w' ? sc > bestScore : sc < bestScore) { bestScore = sc; best = m; }
    }
    const files = 'abcdefgh';
    sel = { r: best.r, c: best.c };
    legal = legalMoves(board, side).filter(m => m.r === best.r && m.c === best.c);
    render();
    setStatus(`Try ${files[best.c]}${8 - best.r} → ${files[best.nc]}${8 - best.nr}`);
  }

  function engineMove() {
    if (twoPlayer || gameOver) return;
    thinking = true;
    setStatus('Thinking…');
    setTimeout(() => {
      const m = aiMove();
      if (m) {
        board = apply(board, m); lastMove = m;
        turn = turn === 'w' ? 'b' : 'w';
        if (typeof beep === 'function') beep([400], .05);
      }
      thinking = false;
      render(); checkEnd();
    }, 130);
  }

  function render() {
    const files = 'abcdefgh';
    let html = '<div class="chessboard">';
    for (let r = 0; r < 8; r++) for (let c = 0; c < 8; c++) {
      const p = board[r][c];
      const dark = (r + c) % 2 === 1;
      const isSel = sel && sel.r === r && sel.c === c;
      const canGo = legal.some(m => m.nr === r && m.nc === c);
      const isCap = canGo && p !== ' ';
      const wasLast = lastMove && ((lastMove.r === r && lastMove.c === c) ||
                                  (lastMove.nr === r && lastMove.nc === c));
      html += `<div class="sq ${dark ? 'd' : 'l'}${isSel ? ' sel' : ''}${wasLast ? ' last' : ''}"
                    data-r="${r}" data-c="${c}">
        ${canGo ? `<span class="dot${isCap ? ' cap' : ''}"></span>` : ''}
        ${p !== ' ' ? `<span class="pc ${isW(p) ? 'w' : 'b'}">${GLYPH[p]}</span>` : ''}
        ${c === 0 ? `<i class="rk">${8-r}</i>` : ''}
        ${r === 7 ? `<i class="fl">${files[c]}</i>` : ''}
      </div>`;
    }
    html += '</div>';
    const L = LEVELS[level] || LEVELS.normal;
    stage().innerHTML = `
      <div class="chessbar">
        <div class="cbgroup">
          ${Object.entries(LEVELS).map(([k, v]) =>
            `<button class="cbtn ${level === k && !twoPlayer ? 'on' : ''}"
                     data-lvl="${k}">${v.label}</button>`).join('')}
          <button class="cbtn ${twoPlayer ? 'on' : ''}" data-lvl="2p">2 Player</button>
        </div>
        <div class="cbgroup">
          <button class="cbtn" id="cUndo" title="Undo your last move">↶ Undo</button>
          <button class="cbtn" id="cHint" title="Suggest a move">💡 Hint</button>
          <button class="cbtn" id="cFlip" title="Swap sides">⇅ ${playAs === 'w' ? 'White' : 'Black'}</button>
          <button class="cbtn" id="cNew">↻ New</button>
        </div>
      </div>` + html +
      `<div class="ghint">${twoPlayer
        ? 'Two players, one board — white moves first'
        : `You are ${playAs === 'w' ? 'white' : 'black'} · ${L.label} opponent`}</div>`;
    stage().querySelectorAll('.sq').forEach(sq =>
      sq.onclick = () => click(+sq.dataset.r, +sq.dataset.c));

    stage().querySelectorAll('.cbtn[data-lvl]').forEach(b => b.onclick = () => {
      const v = b.dataset.lvl;
      if (v === '2p') { twoPlayer = true; localStorage.setItem('fa.chess.2p', '1'); }
      else {
        twoPlayer = false; level = v;
        localStorage.setItem('fa.chess.2p', '0');
        localStorage.setItem('fa.chess.level', v);
      }
      reset(); render();
      setStatus(twoPlayer ? 'Two-player mode — white to move'
                          : `${LEVELS[level].label} mode — your move`);
    });
    const q = id => stage().querySelector(id);
    if (q('#cNew')) q('#cNew').onclick = () => {
      reset(); render(); setStatus('New game — your move.'); };
    if (q('#cUndo')) q('#cUndo').onclick = undoMove;
    if (q('#cHint')) q('#cHint').onclick = showHint;
    if (q('#cFlip')) q('#cFlip').onclick = () => {
      playAs = playAs === 'w' ? 'b' : 'w';
      localStorage.setItem('fa.chess.side', playAs);
      reset(); render();
      if (playAs === 'b' && !twoPlayer) engineMove();
      setStatus(`You are now ${playAs === 'w' ? 'white' : 'black'}.`);
    };
  }

  function click(r, c) {
    if (thinking || gameOver) return;
    const side = twoPlayer ? turn : playAs;
    if (turn !== side) return;
    const p = board[r][c];
    const move = legal.find(m => m.nr === r && m.nc === c);
    if (move) {
      undoStack.push({ board: board.map(x => x.slice()), turn, lastMove });
      if (undoStack.length > 40) undoStack.shift();
      board = apply(board, move); lastMove = move;
      sel = null; legal = []; turn = turn === 'w' ? 'b' : 'w';
      if (typeof beep === 'function') beep([520], .05);
      render();
      if (checkEnd()) return;
      engineMove();
      return;
    }
    const ownsPiece = p !== ' ' && (side === 'w' ? isW(p) : isB(p));
    if (ownsPiece) {
      sel = { r, c };
      legal = legalMoves(board, side).filter(m => m.r === r && m.c === c);
    } else { sel = null; legal = []; }
    render();
  }

  function checkEnd() {
    const moves = legalMoves(board, turn);
    if (moves.length) {
      setStatus(attacked(board, turn)
        ? (turn === 'w' ? 'Check! Your move.' : 'You have it in check.')
        : (turn === 'w' ? 'Your move.' : 'Thinking…'));
      return false;
    }
    gameOver = true;
    if (attacked(board, turn)) {
      const youWon = twoPlayer ? false : (turn !== playAs);
      setStatus(twoPlayer
        ? `Checkmate — ${turn === 'w' ? 'Black' : 'White'} wins!`
        : (youWon ? 'Checkmate — you win! 🎉' : 'Checkmate. I win this one.'));
      if (youWon && typeof confetti === 'function') confetti(140);
    } else setStatus('Stalemate — draw.');
    return true;
  }

  render();
  setStatus(twoPlayer ? 'Two players — white to move.'
                      : `${LEVELS[level].label} mode. You are ${playAs === 'w' ? 'white' : 'black'}.`);
  if (playAs === 'b' && !twoPlayer) engineMove();
  arcadeCleanup = () => {};
}

/* ════════════════════════════════════════════════════════════════
   5. SPOT THE LIAR — logic deduction using real roster names
   ════════════════════════════════════════════════════════════════ */
function gLiar() {
  const pool = (typeof S !== 'undefined' && S.students && S.students.length >= 4)
    ? S.students.map(s => s.Name)
    : ['Aarav', 'Priya', 'Rohan', 'Meera', 'Kabir', 'Ananya', 'Dev', 'Isha'];

  let round = +(localStorage.getItem('fa.liar.round') || 1);
  let streak = 0;

  function pick(n) {
    const c = [...pool];
    const out = [];
    while (out.length < n && c.length) out.push(c.splice((Math.random()*c.length)|0, 1)[0]);
    while (out.length < n) out.push('Student ' + (out.length + 1));
    return out;
  }

  function build() {
    // Generate a puzzle, then verify it has exactly ONE consistent solution.
    // Without this check a puzzle can be ambiguous, which makes it unfair.
    for (let attempt = 0; attempt < 400; attempt++) {
      const who = pick(4);
      const liar = (Math.random() * 4) | 0;
      const culprit = (Math.random() * 4) | 0;

      const claims = who.map((name, i) => {
        const others = [0,1,2,3].filter(j => j !== i);
        const target = others[(Math.random()*others.length)|0];
        const roll = Math.random();
        if (roll < .34)  return { i, kind:'self',   target:i };
        if (roll < .67)  return { i, kind:'accuse', target };
        return { i, kind:'vouch', target };
      });

      // Does claim c evaluate true when `cul` is the culprit?
      const holds = (c, cul) =>
        c.kind === 'self'   ? cul === c.i
      : c.kind === 'accuse' ? cul === c.target
      :                       cul !== c.target;

      // What each speaker actually SAYS (liar states the opposite of reality)
      const stated = claims.map(c =>
        (c.i === liar) ? !holds(c, culprit) : holds(c, culprit));

      // Uniqueness: how many (culprit, liar) pairs explain these statements?
      let solutions = 0, only = null;
      for (let cul = 0; cul < 4; cul++) {
        for (let li = 0; li < 4; li++) {
          const ok = claims.every((c, i) => {
            const truth = holds(c, cul);
            return (i === li) ? (stated[i] !== truth) : (stated[i] === truth);
          });
          if (ok) { solutions++; only = { cul, li }; }
        }
      }
      if (solutions !== 1 || only.li !== liar || only.cul !== culprit) continue;

      // Render each statement as English
      const said = claims.map((c, i) => {
        const yes = stated[i];
        if (c.kind === 'self')
          return yes ? 'Honestly, it was me.' : 'It was not me.';
        if (c.kind === 'accuse')
          return yes ? `It was ${who[c.target]}.`
                     : `${who[c.target]} had nothing to do with it.`;
        return yes ? `I was with ${who[c.target]} the whole time — they are clean.`
                   : `${who[c.target]} was not where they say they were.`;
      });

      return { who, liar, culprit, stmts: who.map((n, i) => ({ name:n, said:said[i] })) };
    }

    // Extremely unlikely fallback
    const who = pick(4);
    return { who, liar:0, culprit:1,
             stmts: who.map((n, i) => ({ name:n,
               said: i === 0 ? `It was ${who[2]}.`
                   : i === 1 ? 'It was not me.'
                   : i === 2 ? 'It was not me.'
                             : `It was ${who[1]}.` })) };
  }

  let puzzle = build();

  const SCENES = [
    'Someone marked themselves present without scanning in.',
    'The whiteboard was wiped before the lesson was photographed.',
    'A pack of XP points went missing from the leaderboard.',
    'Someone rearranged the house standings board overnight.',
    'The classroom projector remote vanished before period 3.',
  ];
  const scene = SCENES[(round - 1) % SCENES.length];

  function render(revealed) {
    stage().innerHTML = `
      <div class="liar">
        <div class="case">
          <span class="tag">Case #${String(round).padStart(3,'0')}</span>
          <p>${scene}</p>
          <small>Three of them are telling the truth. Exactly one is lying.
                 Work out who the liar is.</small>
        </div>
        <div class="suspects">
          ${puzzle.stmts.map((s, i) => `
            <div class="susp ${revealed
                 ? (i === puzzle.liar ? 'liar' : 'honest') : ''}" data-i="${i}">
              <div class="face">${['🧑‍🎓','👩‍🎓','🧑','👧'][i]}</div>
              <div class="sbody">
                <b>${esc2(s.name)}</b>
                <p>“${esc2(s.said)}”</p>
              </div>
              ${revealed ? `<div class="verdict">${i === puzzle.liar ? '🤥 LIAR' : '✅ true'}</div>` : ''}
            </div>`).join('')}
        </div>
        <div class="liarfoot" id="liarFoot">
          ${revealed ? '' : '<span>Click whoever you think is lying.</span>'}
        </div>
      </div>`;
    if (!revealed) {
      stage().querySelectorAll('.susp').forEach(el =>
        el.onclick = () => guess(+el.dataset.i));
    }
  }

  function guess(i) {
    const right = i === puzzle.liar;
    render(true);
    if (right) {
      streak++; round++;
      localStorage.setItem('fa.liar.round', round);
      saveHigh('liar', streak);
      setStatus(`Correct — ${puzzle.stmts[puzzle.liar].name} was lying. Streak: ${streak}`);
      if (typeof beep === 'function') beep([700, 900, 1150], .09);
      if (typeof confetti === 'function') confetti(45);
    } else {
      setStatus(`Not quite — ${puzzle.stmts[puzzle.liar].name} was the liar. Streak reset.`);
      streak = 0;
      if (typeof beep === 'function') beep([340, 220], .14);
    }
    document.getElementById('liarFoot').innerHTML =
      `<div style="text-align:center">
         <p style="font-size:12.5px;color:var(--muted);margin-bottom:10px">
           The culprit was <b style="color:var(--text)">${esc2(puzzle.stmts[puzzle.culprit].name)}</b>.
         </p>
         <button class="btn primary" id="nextCase">Next case →</button>
       </div>`;
    document.getElementById('nextCase').onclick = () => {
      puzzle = build(); render(false);
      setStatus('Who is lying this time?');
    };
  }

  render(false);
  setStatus('Read the statements. Who is lying?');
  arcadeCleanup = () => {};
}

const esc2 = v => String(v ?? '').replace(/[&<>"']/g, c =>
  ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));

function gDungeon() {
  const W = 720, H = 380, GROUND = H - 40, GRAV = 0.55;
  stage().innerHTML = `<canvas id="cv" width="${W}" height="${H}" tabindex="0"></canvas>
    <div class="ghint">WASD/Space move &amp; jump · Z melee · X shoot (aim with mouse) · P pause · R restart</div>`;
  const cv = document.getElementById('cv'), cx = cv.getContext('2d');
  cv.style.cursor = 'crosshair';

  const plats = [
    { x:0, y:GROUND, w:W, h:40 },
    { x:80, y:280, w:140, h:16 },
    { x:300, y:220, w:140, h:16 },
    { x:520, y:280, w:140, h:16 },
    { x:220, y:150, w:120, h:16 },
  ];

  let player, bolts, mobs, particles, gold, kills, wave, hp, maxHp, over, mouse, lastShot, lastMelee;
  let waveQueued = false, paused = false;

  function newGame() {
    player = { x:60, y:GROUND-40, w:26, h:38, vx:0, vy:0, onGround:false, facing:1 };
    bolts = []; mobs = []; particles = [];
    gold = 0; kills = 0; wave = 1; hp = 10; maxHp = 10; over = false;
    waveQueued = false;
    lastShot = 0; lastMelee = 0;
    mouse = { x:400, y:200 };
    spawnWave();
  }

  function spawnWave() {
    const n = 2 + Math.floor(wave * 1.4);
    for (let i = 0; i < n; i++) {
      const fromLeft = Math.random() < 0.5;
      mobs.push({
        x: fromLeft ? -30 - Math.random()*80 : W + 30 + Math.random()*80,
        y: GROUND - 30, w:24, h:30, vy:0,
        hp: 2 + Math.floor(wave * 0.7), maxHp: 2 + Math.floor(wave * 0.7),
        speed: 0.9 + wave * 0.08 + Math.random()*0.3,
        hitCd: 0,
      });
    }
  }

  const held = new Set();
  const kd = e => {
    const k = e.key.toLowerCase();
    if (['w','a','s','d',' ','z','x','r','p'].includes(k)) e.preventDefault();
    held.add(k);
    if (k === 'r') { held.clear(); newGame(); return; }   // FIX: restart any time
    if (k === 'p') { paused = !paused; setStatus(paused ? 'Paused — press P' : 'Resumed'); return; }
    if (k === 'z') melee();
    if (k === 'x') shoot();
  };
  const ku = e => held.delete(e.key.toLowerCase());
  window.addEventListener('keydown', kd);
  window.addEventListener('keyup', ku);
  const mv = e => { const r = cv.getBoundingClientRect(); mouse.x = e.clientX - r.left; mouse.y = e.clientY - r.top; };
  cv.addEventListener('mousemove', mv);
  const mc = () => shoot();
  cv.addEventListener('mousedown', mc);

  function melee() {
    if (over || performance.now() - lastMelee < 260) return;
    lastMelee = performance.now();
    beep([420, 300], .06);
    const reach = 34;
    const hitBox = { x: player.x + (player.facing > 0 ? player.w : -reach), y: player.y, w: reach, h: player.h };
    mobs.forEach(m => {
      if (m.x < hitBox.x + hitBox.w && m.x + m.w > hitBox.x && m.y < hitBox.y + hitBox.h && m.y + m.h > hitBox.y) {
        m.hp -= 2; m.hitCd = 8; spark(m.x + m.w/2, m.y + m.h/2);
      }
    });
    cleanupMobs();
  }

  function shoot() {
    if (over || performance.now() - lastShot < 320) return;
    lastShot = performance.now();
    beep([880, 1180], .05);
    const cx0 = player.x + player.w/2, cy0 = player.y + player.h/2;
    const ang = Math.atan2(mouse.y - cy0, mouse.x - cx0);
    bolts.push({ x:cx0, y:cy0, vx:Math.cos(ang)*8.5, vy:Math.sin(ang)*8.5 });
  }

  function spark(x, y) {
    for (let i = 0; i < 6; i++) particles.push({ x, y, vx:(Math.random()-0.5)*4, vy:(Math.random()-1.2)*3, life:18, col:'#ff9f5d' });
  }

  function cleanupMobs() {
    const before = mobs.length;
    mobs = mobs.filter(m => m.hp > 0 || (spark(m.x+m.w/2, m.y+m.h/2), false));
    const dead = before - mobs.length;
    if (dead) {
      kills += dead; gold += dead * (3 + wave);
      // FIX: cleanupMobs() is called from melee, bolts and step in the same
      // frame. Without this guard an empty board could advance the wave
      // more than once and queue several spawns at the same time.
      if (mobs.length === 0 && !waveQueued) {
        waveQueued = true;
        wave++; hp = Math.min(maxHp, hp + 1);
        setTimeout(() => {
          waveQueued = false;
          if (!over) spawnWave();
        }, 700);
      }
    }
  }

  function step() {
    if (over || paused) return;
    // movement
    const speed = 3.6;
    if (held.has('a')) { player.vx = -speed; player.facing = -1; }
    else if (held.has('d')) { player.vx = speed; player.facing = 1; }
    else player.vx *= 0.7;
    if ((held.has('w') || held.has(' ')) && player.onGround) { player.vy = -10.4; player.onGround = false; beep([560], .04); }

    player.vy += GRAV;
    player.x += player.vx; player.y += player.vy;
    player.x = Math.max(0, Math.min(W - player.w, player.x));

    player.onGround = false;
    plats.forEach(p => {
      // FIX: the old test only caught a landing if the feet ended up inside
      // a 14px band. Falling faster than that (easy after a jump) tunnelled
      // straight through. Now we check whether the feet CROSSED the top of
      // the platform during this frame, which cannot be skipped.
      const feetNow = player.y + player.h;
      const feetPrev = feetNow - player.vy;
      const overlapX = player.x + player.w > p.x && player.x < p.x + p.w;
      if (overlapX && player.vy >= 0 && feetPrev <= p.y + 1 && feetNow >= p.y) {
        player.y = p.y - player.h; player.vy = 0; player.onGround = true;
      }
    });
    if (player.y + player.h > GROUND) { player.y = GROUND - player.h; player.vy = 0; player.onGround = true; }

    // bolts
    bolts.forEach(b => { b.x += b.vx; b.y += b.vy; });
    bolts = bolts.filter(b => b.x > -10 && b.x < W + 10 && b.y > -10 && b.y < H + 10);
    bolts.forEach(b => {
      mobs.forEach(m => {
        if (b.hit) return;
        if (b.x > m.x && b.x < m.x + m.w && b.y > m.y && b.y < m.y + m.h) {
          m.hp -= 1; b.hit = true; spark(b.x, b.y);
        }
      });
    });
    bolts = bolts.filter(b => !b.hit);
    cleanupMobs();

    // mobs chase + gravity
    mobs.forEach(m => {
      m.vy = (m.vy || 0) + GRAV;
      m.y += m.vy;
      if (m.y + m.h > GROUND) { m.y = GROUND - m.h; m.vy = 0; }
      const dx = (player.x + player.w/2) - (m.x + m.w/2);
      m.x += Math.sign(dx) * m.speed;
      if (m.hitCd > 0) m.hitCd--;
      // FIX: this compared player.y to m.y directly, but the sprites are
      // different heights (38 vs 30), so mobs could hit through a floor or
      // miss when clearly touching. Use a real box overlap instead.
      const touching = player.x < m.x + m.w && player.x + player.w > m.x &&
                       player.y < m.y + m.h && player.y + player.h > m.y;
      if (touching && m.hitCd === 0) {
        m.hitCd = 40; hp -= 1; beep([260, 180], .1);
        if (hp <= 0) { over = true; const rec = saveHigh('dungeon', kills*10+gold); setStatus(`You fell — wave ${wave}, ${kills} kills, ${gold} gold${rec?' · new best!':''}. Press R to retry.`); beep([220,160,120], .16); }
      }
    });

    particles.forEach(p => { p.x += p.vx; p.y += p.vy; p.vy += 0.15; p.life--; });
    particles = particles.filter(p => p.life > 0);

    if (!over) setStatus(`Wave ${wave} · HP ${hp}/${maxHp} · ${kills} kills · ${gold}g — Z melee, X shoot`);
  }

  function draw() {
    const grd = cx.createLinearGradient(0, 0, 0, H);
    grd.addColorStop(0, '#0e0a1a'); grd.addColorStop(1, '#161022');
    cx.fillStyle = grd; cx.fillRect(0, 0, W, H);

    cx.fillStyle = '#241b38';
    plats.forEach(p => { rr(cx, p.x, p.y, p.w, p.h, 4); cx.fill(); });

    // player
    cx.fillStyle = '#5b8cff'; rr(cx, player.x, player.y, player.w, player.h, 5); cx.fill();
    cx.fillStyle = '#eef2fb';
    cx.beginPath(); cx.arc(player.x + player.w/2 + player.facing*4, player.y + 10, 3, 0, 7); cx.fill();

    // aim line
    if (!over) {
      cx.strokeStyle = 'rgba(255,255,255,.15)'; cx.setLineDash([3,5]);
      cx.beginPath(); cx.moveTo(player.x+player.w/2, player.y+player.h/2); cx.lineTo(mouse.x, mouse.y); cx.stroke();
      cx.setLineDash([]);
    }

    // bolts
    cx.fillStyle = '#ffd166';
    bolts.forEach(b => { cx.beginPath(); cx.arc(b.x, b.y, 3.4, 0, 7); cx.fill(); });

    // mobs
    mobs.forEach(m => {
      cx.fillStyle = '#c86bff'; rr(cx, m.x, m.y, m.w, m.h, 5); cx.fill();
      cx.fillStyle = 'rgba(0,0,0,.5)'; cx.fillRect(m.x, m.y - 8, m.w, 4);
      cx.fillStyle = '#ff5d73'; cx.fillRect(m.x, m.y - 8, m.w * (m.hp / m.maxHp), 4);
    });

    particles.forEach(p => { cx.globalAlpha = Math.max(0, p.life/18); cx.fillStyle = p.col; cx.fillRect(p.x, p.y, 3, 3); });
    cx.globalAlpha = 1;

    if (over) {
      cx.fillStyle = 'rgba(5,7,13,.78)'; cx.fillRect(0, 0, W, H);
      cx.fillStyle = '#eef2fb'; cx.font = '700 26px system-ui'; cx.textAlign = 'center';
      cx.fillText('You fell', W/2, H/2 - 8);
      cx.font = '400 14px system-ui'; cx.fillStyle = '#94a3bb';
      cx.fillText(`Wave ${wave} · ${kills} kills · ${gold} gold — Press R to retry`, W/2, H/2 + 20);
    }
  }

  // FIX: newGame() must run BEFORE the loop starts. It creates `player`,
  // and step() dereferences player.vx on its very first frame — starting
  // the loop first crashed the game instantly with
  // "Cannot read properties of undefined (reading 'vx')".
  let raf;
  newGame();
  cv.focus();
  (function loop() { raf = requestAnimationFrame(loop); step(); draw(); })();

  arcadeCleanup = () => {
    cancelAnimationFrame(raf);
    window.removeEventListener('keydown', kd);
    window.removeEventListener('keyup', ku);
    window.removeEventListener('blur', blur);
    cv.removeEventListener('mousemove', mv);
    cv.removeEventListener('mousedown', mc);
  };
}

function gPong() {
  const W = 520, H = 320, PW = 10, PH = 68, BALL = 8, WIN = 7;
  stage().innerHTML = `<canvas id="cv" width="${W}" height="${H}"></canvas>
    <div class="ghint">Mouse or Arrow keys / W · S — first to ${WIN}</div>`;
  const cv = document.getElementById('cv'), cx = cv.getContext('2d');

  let py = H / 2 - PH / 2, ay = H / 2 - PH / 2;
  let bx = W / 2, by = H / 2, bvx = 4.4, bvy = 2.1;
  let ps = 0, as = 0, over = false;

  function resetBall(dir) {
    bx = W / 2; by = H / 2;
    bvx = dir * 4.4; bvy = (Math.random() * 3.2 - 1.6);
  }

  const move = e => {
    const r = cv.getBoundingClientRect();
    py = e.clientY - r.top - PH / 2;
  };
  cv.addEventListener('mousemove', move);

  const held = new Set();
  const kd = e => {
    const k = e.key.toLowerCase();
    if (['arrowup','arrowdown','w','s'].includes(k)) e.preventDefault();
    held.add(k);
    if (k === 'r' && over) { ps = 0; as = 0; over = false; resetBall(1); setStatus('0 – 0'); }
  };
  const ku = e => held.delete(e.key.toLowerCase());
  // FIX: alt-tabbing while holding a key left it "held" forever, so the
  // player kept running on its own after coming back.
  const blur = () => held.clear();
  window.addEventListener('keydown', kd);
  window.addEventListener('keyup', ku);
  window.addEventListener('blur', blur);

  function step() {
    if (over) return;
    if (held.has('arrowup') || held.has('w')) py -= 6.2;
    if (held.has('arrowdown') || held.has('s')) py += 6.2;
    py = Math.max(0, Math.min(H - PH, py));

    const target = by - PH / 2 + (Math.random() * 26 - 13);
    ay += Math.max(-5.2, Math.min(5.2, (target - ay) * 0.11));
    ay = Math.max(0, Math.min(H - PH, ay));

    bx += bvx; by += bvy;
    if (by < BALL || by > H - BALL) { bvy *= -1; beep([560], .04); }

    if (bx - BALL < PW + 4 && by > py && by < py + PH && bvx < 0) {
      bvx *= -1.05; bvy += (by - (py + PH / 2)) * 0.06;
      beep([620, 780], .05);
    }
    if (bx + BALL > W - PW - 4 && by > ay && by < ay + PH && bvx > 0) {
      bvx *= -1.05; bvy += (by - (ay + PH / 2)) * 0.06;
      beep([620, 780], .05);
    }

    if (bx < 0) { as++; setStatus(`${ps} – ${as}`); check(); if (!over) resetBall(1); }
    if (bx > W) { ps++; setStatus(`${ps} – ${as}`); check(); if (!over) resetBall(-1); }
  }

  function check() {
    if (ps >= WIN || as >= WIN) {
      over = true;
      const won = ps > as;
      const rec = won ? saveHigh('pong', ps) : false;
      setStatus(`${won ? 'You win' : 'AI wins'} ${ps} – ${as}${rec ? ' · new best!' : ''} — Press R to play again`);
      if (won) { beep([660, 880, 1100], .09); if (rec) confetti(35); }
      else beep([300, 220], .13);
    }
  }

  function draw() {
    cx.fillStyle = '#080b12'; cx.fillRect(0, 0, W, H);
    cx.strokeStyle = 'rgba(255,255,255,.08)';
    cx.setLineDash([6, 8]);
    cx.beginPath(); cx.moveTo(W / 2, 0); cx.lineTo(W / 2, H); cx.stroke();
    cx.setLineDash([]);
    cx.fillStyle = '#5b8cff'; rr(cx, 4, py, PW, PH, 4); cx.fill();
    cx.fillStyle = '#ff5d73'; rr(cx, W - PW - 4, ay, PW, PH, 4); cx.fill();
    cx.fillStyle = '#eef2fb'; cx.beginPath(); cx.arc(bx, by, BALL, 0, 7); cx.fill();
  }

  let raf;
  (function loop() { raf = requestAnimationFrame(loop); step(); draw(); })();
  setStatus('0 – 0');
  arcadeCleanup = () => {
    cancelAnimationFrame(raf);
    cv.removeEventListener('mousemove', move);
    window.removeEventListener('keydown', kd);
    window.removeEventListener('keyup', ku);
  };
}

function gTtt() {
  stage().innerHTML = `<div class="ttt" id="tttGrid"></div>
    <div class="ghint">You are X · click a square</div>`;
  const wrap = document.getElementById('tttGrid');
  wrap.style.cssText = 'display:grid;grid-template-columns:repeat(3,84px);gap:8px;justify-content:center';

  let board, turn, done;

  function newGame() { board = Array(9).fill(''); turn = 'X'; done = false; render(); setStatus('Your move'); }

  const LINES = [[0,1,2],[3,4,5],[6,7,8],[0,3,6],[1,4,7],[2,5,8],[0,4,8],[2,4,6]];
  function winner(b) {
    for (const [a, c, d] of LINES) if (b[a] && b[a] === b[c] && b[a] === b[d]) return b[a];
    return b.includes('') ? null : 'draw';
  }

  function minimax(b, player) {
    const w = winner(b);
    if (w === 'O') return { score: 1 };
    if (w === 'X') return { score: -1 };
    if (w === 'draw') return { score: 0 };
    const moves = [];
    for (let i = 0; i < 9; i++) if (!b[i]) {
      b[i] = player;
      const r = minimax(b, player === 'O' ? 'X' : 'O');
      moves.push({ i, score: r.score });
      b[i] = '';
    }
    moves.sort((a, c) => player === 'O' ? c.score - a.score : a.score - c.score);
    return moves[0];
  }

  function aiMove() {
    if (done) return;
    const best = minimax(board.slice(), 'O');
    if (best && best.i !== undefined) board[best.i] = 'O';
    afterMove();
  }

  function afterMove() {
    const w = winner(board);
    render();
    if (w) {
      done = true;
      if (w === 'draw') { setStatus("It's a draw — press R to play again"); beep([440, 370], .09); }
      else if (w === 'X') { setStatus('You win! Press R to play again'); saveHigh('ttt', 1); confetti(35); beep([660, 880, 1100], .09); }
      else { setStatus('AI wins — press R to try again'); beep([300, 220], .13); }
      return;
    }
    turn = turn === 'X' ? 'O' : 'X';
    setStatus(turn === 'X' ? 'Your move' : 'AI thinking…');
    if (turn === 'O') setTimeout(aiMove, 260);
  }

  function render() {
    wrap.innerHTML = board.map((v, i) => `
      <div data-i="${i}" style="width:84px;height:84px;border-radius:12px;background:rgba(255,255,255,.045);
        display:flex;align-items:center;justify-content:center;font-size:34px;font-weight:800;cursor:pointer;
        color:${v === 'X' ? '#5b8cff' : '#ff5d73'}">${v}</div>`).join('');
    wrap.querySelectorAll('[data-i]').forEach(el => el.onclick = () => {
      const i = +el.dataset.i;
      if (done || turn !== 'X' || board[i]) return;
      board[i] = 'X';
      afterMove();
    });
  }

  const key = e => { if (e.key.toLowerCase() === 'r') newGame(); };
  window.addEventListener('keydown', key);
  newGame();
  arcadeCleanup = () => window.removeEventListener('keydown', key);
}

function gMemory() {
  const ICONS = ['🐍','🦖','👻','♟','🕵','🏓','🔢','🧠','🎯','🎲','🚀','⭐'];
  const pairs = ICONS.slice(0, 8);
  let deck = [...pairs, ...pairs].sort(() => Math.random() - 0.5)
    .map((ic, i) => ({ id: i, ic, open: false, matched: false }));
  let openIds = [], moves = 0, matched = 0, lock = false, start = Date.now(), timer;

  stage().innerHTML = `<div class="memgrid" id="memGrid"></div>
    <div class="ghint">Find all ${pairs.length} pairs · fewer moves is better</div>`;
  const wrap = document.getElementById('memGrid');
  wrap.style.cssText = 'display:grid;grid-template-columns:repeat(4,72px);gap:8px;justify-content:center';

  function render() {
    wrap.innerHTML = deck.map(c => `
      <div data-id="${c.id}" style="width:72px;height:72px;border-radius:10px;cursor:pointer;
        display:flex;align-items:center;justify-content:center;font-size:30px;
        background:${c.matched ? 'rgba(46,230,168,.15)' : c.open ? 'rgba(91,140,255,.18)' : 'rgba(255,255,255,.05)'};
        transition:background .15s">${c.open || c.matched ? c.ic : ''}</div>`).join('');
    wrap.querySelectorAll('[data-id]').forEach(el => el.onclick = () => flip(+el.dataset.id));
  }

  function flip(id) {
    if (lock) return;
    const card = deck[id];
    if (card.open || card.matched) return;
    card.open = true; openIds.push(id); render();
    beep([560], .04);
    if (openIds.length === 2) {
      moves++; lock = true;
      const [a, b] = openIds.map(i => deck[i]);
      if (a.ic === b.ic) {
        a.matched = b.matched = true; matched++;
        openIds = []; lock = false; render();
        beep([740, 980], .06);
        if (matched === pairs.length) finish();
      } else {
        setTimeout(() => { a.open = b.open = false; openIds = []; lock = false; render(); }, 550);
      }
    }
  }

  function finish() {
    clearInterval(timer);
    const secs = Math.round((Date.now() - start) / 1000);
    const rec = saveHigh('memory', 100000 - (moves * 1000 + secs)); // fewer moves+time = higher score
    setStatus(`Cleared in ${moves} moves, ${secs}s${rec ? ' · new best!' : ''} — Press R to play again`);
    confetti(40); beep([660, 880, 1100], .09);
  }

  function tick() {
    if (matched === pairs.length) return;
    setStatus(`Moves: ${moves} · ${Math.round((Date.now() - start) / 1000)}s`);
  }
  timer = setInterval(tick, 1000);

  const key = e => {
    if (e.key.toLowerCase() !== 'r') return;
    clearInterval(timer);
    deck = [...pairs, ...pairs].sort(() => Math.random() - 0.5).map((ic, i) => ({ id: i, ic, open: false, matched: false }));
    openIds = []; moves = 0; matched = 0; lock = false; start = Date.now();
    timer = setInterval(tick, 1000);
    render(); setStatus('Moves: 0 · 0s');
  };
  window.addEventListener('keydown', key);

  render(); setStatus('Moves: 0 · 0s');
  arcadeCleanup = () => { clearInterval(timer); window.removeEventListener('keydown', key); };
}

/* ════════════════════════════════════════════════════════════════
   10. 2048
   ════════════════════════════════════════════════════════════════ */
function g2048() {
  const N = 4, CELL = 96, GAP = 10, PAD = 12;
  const SIZE = PAD * 2 + N * CELL + (N - 1) * GAP;

  stage().innerHTML = `<canvas id="cv" width="${SIZE}" height="${SIZE}"></canvas>
    <div class="ghint">Arrow keys or WASD · <kbd>U</kbd> undo · <kbd>R</kbd> restart</div>`;
  const cv = document.getElementById('cv'), cx = cv.getContext('2d');

  const TILE = {
    2:['#eee4da','#776e65'],      4:['#ede0c8','#776e65'],
    8:['#f2b179','#f9f6f2'],     16:['#f59563','#f9f6f2'],
    32:['#f67c5f','#f9f6f2'],    64:['#f65e3b','#f9f6f2'],
    128:['#edcf72','#f9f6f2'],  256:['#edcc61','#f9f6f2'],
    512:['#edc850','#f9f6f2'], 1024:['#edc53f','#f9f6f2'],
    2048:['#edc22e','#f9f6f2'], 4096:['#3c3a32','#f9f6f2'],
  };

  let grid, score, best, over, won, history;

  function reset() {
    grid = Array.from({ length: N }, () => Array(N).fill(0));
    score = 0; over = false; won = false; history = [];
    best = +(localStorage.getItem('fa.hs.2048') || 0);
    addTile(); addTile();
    setStatus('Score 0');
  }

  function addTile() {
    const free = [];
    for (let r = 0; r < N; r++) for (let c = 0; c < N; c++)
      if (!grid[r][c]) free.push([r, c]);
    if (!free.length) return;
    const [r, c] = free[(Math.random() * free.length) | 0];
    grid[r][c] = Math.random() < 0.9 ? 2 : 4;
  }

  // Collapse one row to the left; returns the new row and points gained.
  function slide(row) {
    const vals = row.filter(v => v);
    let gained = 0;
    for (let i = 0; i < vals.length - 1; i++) {
      if (vals[i] === vals[i + 1]) {
        vals[i] *= 2;
        gained += vals[i];
        vals.splice(i + 1, 1);
      }
    }
    while (vals.length < N) vals.push(0);
    return [vals, gained];
  }

  const rotate = g => g[0].map((_, i) => g.map(r => r[i]).reverse());   // 90° CW

  function move(dir) {
    if (over) return;
    const before = JSON.stringify(grid);
    history.push({ grid: JSON.parse(before), score });
    if (history.length > 20) history.shift();

    // Rotate so every direction becomes a leftward slide, then rotate back.
    let turns = { left: 0, up: 1, right: 2, down: 3 }[dir];
    for (let i = 0; i < turns; i++) grid = rotate(grid);

    let gained = 0;
    grid = grid.map(row => {
      const [next, g] = slide(row);
      gained += g;
      return next;
    });

    for (let i = 0; i < (4 - turns) % 4; i++) grid = rotate(grid);

    if (JSON.stringify(grid) === before) {
      history.pop();                      // nothing moved; not a real turn
      return;
    }

    score += gained;
    if (gained && typeof beep === 'function') beep([440 + Math.min(600, gained)], .05);
    addTile();

    if (!won && grid.some(r => r.some(v => v >= 2048))) {
      won = true;
      setStatus(`2048! Score ${score} — keep going for a bigger tile.`);
      if (typeof confetti === 'function') confetti(120);
      if (typeof beep === 'function') beep([660, 880, 1100], .12);
    } else {
      setStatus(`Score ${score}`);
    }

    if (!canMove()) {
      over = true;
      const rec = score > best;
      if (rec) { best = score; localStorage.setItem('fa.hs.2048', score); }
      if (typeof saveHigh === 'function') saveHigh('2048', score);
      setStatus(`No moves left — ${score} points${rec ? ' · new best!' : ''}. Press R.`);
    }
  }

  function canMove() {
    for (let r = 0; r < N; r++) for (let c = 0; c < N; c++) {
      if (!grid[r][c]) return true;
      if (c < N - 1 && grid[r][c] === grid[r][c + 1]) return true;
      if (r < N - 1 && grid[r][c] === grid[r + 1][c]) return true;
    }
    return false;
  }

  function undo() {
    if (!history.length) return setStatus('Nothing to undo.');
    const snap = history.pop();
    grid = snap.grid; score = snap.score; over = false;
    setStatus(`Score ${score} (undone)`);
  }

  const key = e => {
    const k = e.key.toLowerCase();
    const dir = { arrowleft:'left', a:'left', arrowright:'right', d:'right',
                  arrowup:'up', w:'up', arrowdown:'down', s:'down' }[k];
    if (dir) { e.preventDefault(); move(dir); return; }
    if (k === 'r') { e.preventDefault(); reset(); }
    if (k === 'u') { e.preventDefault(); undo(); }
  };
  window.addEventListener('keydown', key);

  // swipe support, so it works on a touchscreen too
  let sx = 0, sy = 0;
  const ts = e => { const t = e.touches[0]; sx = t.clientX; sy = t.clientY; };
  const te = e => {
    const t = e.changedTouches[0];
    const dx = t.clientX - sx, dy = t.clientY - sy;
    if (Math.hypot(dx, dy) < 24) return;
    move(Math.abs(dx) > Math.abs(dy)
      ? (dx > 0 ? 'right' : 'left')
      : (dy > 0 ? 'down' : 'up'));
  };
  cv.addEventListener('touchstart', ts, { passive: true });
  cv.addEventListener('touchend', te, { passive: true });

  function draw() {
    cx.fillStyle = '#161c28';
    rr(cx, 0, 0, SIZE, SIZE, 16); cx.fill();

    for (let r = 0; r < N; r++) for (let c = 0; c < N; c++) {
      const x = PAD + c * (CELL + GAP), y = PAD + r * (CELL + GAP);
      cx.fillStyle = 'rgba(255,255,255,.045)';
      rr(cx, x, y, CELL, CELL, 9); cx.fill();

      const v = grid[r][c];
      if (!v) continue;
      const [bg, fg] = TILE[v] || TILE[4096];
      cx.fillStyle = bg;
      rr(cx, x, y, CELL, CELL, 9); cx.fill();
      cx.fillStyle = fg;
      cx.textAlign = 'center';
      cx.textBaseline = 'middle';
      const size = v < 100 ? 40 : v < 1000 ? 33 : v < 10000 ? 26 : 21;
      cx.font = `800 ${size}px system-ui,sans-serif`;
      cx.fillText(String(v), x + CELL / 2, y + CELL / 2 + 1);
    }

    if (over) {
      cx.fillStyle = 'rgba(11,14,20,.80)';
      rr(cx, 0, 0, SIZE, SIZE, 16); cx.fill();
      cx.fillStyle = '#eef2fb'; cx.textAlign = 'center';
      cx.font = '800 30px system-ui';
      cx.fillText('No moves left', SIZE / 2, SIZE / 2 - 10);
      cx.font = '400 14px system-ui'; cx.fillStyle = '#94a3bb';
      cx.fillText(`${score} points · best ${best} · press R`, SIZE / 2, SIZE / 2 + 22);
    }
  }

  let raf;
  reset();
  (function loop() { raf = requestAnimationFrame(loop); draw(); })();

  arcadeCleanup = () => {
    cancelAnimationFrame(raf);
    window.removeEventListener('keydown', key);
    cv.removeEventListener('touchstart', ts);
    cv.removeEventListener('touchend', te);
  };
}

window.openArcade = openArcade;
window.closeArcade = closeArcade;
window.showArcadeMenu = showArcadeMenu;
"""

INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FaceAttend · Smart Attendance</title>
<link rel="icon" href="/static/icon.svg" type="image/svg+xml">
<link rel="stylesheet" href="/static/app.css">
</head>
<body>

<!-- ══ SIGN IN ══ -->
<div class="gate" id="gate">
  <div class="gatecard">
    <div class="gatebrand">
      <div class="gatelogo">FA</div>
      <div>
        <h1 class="gatetitle">FaceAttend</h1>
        <p class="gatesub" id="gateSub">Sign in to continue</p>
      </div>
    </div>

    <div id="gatePicker" style="display:none">
      <div class="pickgrid" id="pickGrid"></div>
      <button type="button" class="pickother" id="pickOther">Use another account</button>
    </div>

    <form id="gateForm" autocomplete="on">
      <button type="button" class="pickback" id="pickBack" style="display:none">
        ← Back to accounts
      </button>
      <div class="gatefield" id="gateDisplayWrap" style="display:none">
        <label for="gateDisplay">Your name</label>
        <input id="gateDisplay" placeholder="Ms Sharma" autocomplete="name">
      </div>
      <div class="gatefield">
        <label for="gateUser">Username</label>
        <input id="gateUser" placeholder="teacher" autocomplete="username" required>
      </div>
      <div class="gatefield">
        <label for="gatePass">Password</label>
        <div class="gatepass">
          <input id="gatePass" type="password" placeholder="••••••••"
                 autocomplete="current-password" required>
          <button type="button" class="gateeye" id="gateEye" title="Show password">👁</button>
        </div>
        <div class="gatemeter" id="gateMeter" style="display:none"><i></i></div>
      </div>
      <div class="gatefield" id="gateConfirmWrap" style="display:none">
        <label for="gateConfirm">Confirm password</label>
        <input id="gateConfirm" type="password" placeholder="••••••••"
               autocomplete="new-password">
      </div>

      <div class="gateerr" id="gateErr"></div>
      <button type="submit" class="btn primary block" id="gateGo"
              style="margin-top:4px;height:44px">Sign in</button>
    </form>

    <div class="gateswap">
      <span id="gateSwapText">First time here?</span>
      <button type="button" id="gateSwap">Create an account</button>
    </div>

    <p class="gatenote">
      Accounts are stored on this computer only. Passwords are hashed with
      PBKDF2 and never saved as plain text.
    </p>
  </div>
</div>

<div class="aurora"><b></b><b></b><b></b></div>

<div class="shell" id="shell">

  <!-- ══ NAV RAIL ══ -->
  <nav class="rail">
    <div class="brand" id="brand" title="FaceAttend">FA</div>

    <button class="navbtn on" data-page="home">⌂<span class="tip">Home &amp; Scanner</span></button>
    <button class="navbtn" data-page="analytics">📊<span class="tip">Analytics</span></button>
    <button class="navbtn" data-page="bot">🤖<span class="tip">AI Buddy</span></button>
    <button class="navbtn" data-page="mail">✉<span class="tip">Messages</span></button>
    <button class="navbtn" data-page="awards">🏅<span class="tip">Badges &amp; Certificates</span></button>
    <button class="navbtn" data-page="console">📋<span class="tip">Console</span></button>
    <button class="navbtn" data-page="notes">📝<span class="tip">Notepad</span></button>
    <button class="navbtn" data-page="settings">⚙<span class="tip">Settings</span></button>

    <div class="rail-foot">
      <div class="dot" id="scanDot" title="Scanner idle"></div>
    </div>
  </nav>

  <!-- ══ SIDE PANEL ══ -->
  <aside class="panel" id="panel">

    <!-- HOME -->
    <section class="page on" data-page="home">
      <div class="ptitle">FaceAttend <span style="font-size:11px;color:var(--muted);font-weight:600" id="verLabel"></span></div>
      <div class="psub">Smart attendance, scanning &amp; rewards</div>

      <div class="cam" id="cam">
        <div class="off" id="camOff">
          <div class="ic">📷</div>
          <div style="font-size:13px;font-weight:650">Camera is off</div>
          <div style="font-size:11.5px;margin-top:4px">Press Start Scan to begin</div>
        </div>
        <img id="camFeed" style="display:none" alt="Live camera">
        <div class="badge" style="display:none" id="camBadge">
          <span class="dot live"></span> LIVE
        </div>
              </div>

      <div id="windowNow" style="margin-top:10px;padding:10px 13px;
           border-radius:12px;background:var(--surface);
           border:1px solid var(--stroke);font-size:11.5px;display:none">
        <div style="display:flex;align-items:center;gap:8px">
          <span id="windowDot" style="width:8px;height:8px;border-radius:50%;
                background:var(--faint);flex:none"></span>
          <b id="windowLabel" style="font-size:12.5px"></b>
        </div>
        <div id="windowSub" style="color:var(--muted);margin-top:4px"></div>
      </div>

      <div id="recogWarn" style="display:none;margin-top:10px;padding:11px 13px;
           border-radius:12px;background:rgba(255,176,32,.12);
           border:1px solid var(--warn);font-size:11.5px;line-height:1.55">
        <b style="color:var(--warn)">Automatic marking is off</b>
        <div id="recogWarnMsg" style="color:var(--muted);margin-top:4px"></div>
      </div>

      <div class="row" style="margin-top:12px">
        <button class="btn primary" id="btnScan" style="flex:1">▶ Start Scan</button>
        <button class="btn" id="btnRefresh">↻</button>
      </div>

      <div class="recog" id="recog">
        <div class="lb">Last recognition</div>
        <h3 id="recogName">Waiting for student…</h3>
        <div class="meta" id="recogMeta">Start scanning to mark attendance.</div>
      </div>

      <div class="sec">Activity feed</div>
      <div class="feed" id="feed"></div>
    </section>

    <!-- ANALYTICS -->
    <section class="page" data-page="analytics">
      <div class="wide-head">
        <div><div class="ptitle">📊 Analytics</div>
          <div class="psub">Attendance trends, top students, XP leaders</div></div>
      </div>
      <div class="pane-switch">
        <span class="pane-label">View</span>
        <button class="pane-btn on" id="analyticsTab">📊 Analytics</button>
        <button class="pane-btn" id="housesTab">🏠 Houses</button>
      </div>

      <div id="analyticsPane" class="pane-body">
        <div class="analytics-grid">
          <div class="analytics-main">
            <div class="grid2" style="margin-bottom:14px">
              <div class="card" style="padding:14px">
                <div class="lb" style="font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.9px;font-weight:750">Avg attendance</div>
                <div style="font-size:26px;font-weight:850;color:var(--ok);letter-spacing:-1px" id="avgAtt">—</div>
              </div>
              <div class="card" style="padding:14px">
                <div class="lb" style="font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.9px;font-weight:750">School days</div>
                <div style="font-size:26px;font-weight:850;letter-spacing:-1px" id="schoolDays">—</div>
              </div>
            </div>
            <div class="card">
              <div class="sec" style="margin-top:0">Attendance by house</div>
              <div class="bars" id="houseBars"></div>
            </div>
          </div>
          <div class="analytics-side">
            <div class="sec" style="margin-top:0">30-day trend</div>
            <div class="card" style="padding:14px">
              <svg class="chart" id="trend" viewBox="0 0 300 130" preserveAspectRatio="none"></svg>
            </div>
            <div class="sec">Top students by XP</div>
            <div class="lb" id="topXp"></div>
          </div>
        </div>
      </div>

      <div id="housesPane" class="pane-body" style="display:none">
        <div class="houses" id="houseCards"></div>
        <div class="sec">Ranking</div>
        <div class="lb" id="houseRank"></div>
      </div>
    </section>

    <!-- BOT -->
    <section class="page" data-page="bot">
      <div class="bothead">
        <div class="r1">
          <div class="avatar">🤖</div>
          <div style="flex:1;min-width:0">
            <h3>AI Buddy</h3>
            <div class="status" id="botStatus">
              <span class="d"></span><span>Reading live roster · <b>offline brain</b></span>
            </div>
          </div>
        </div>
        <p>Ask about attendance, students or XP — or tell me to run things for you.
           I only answer from your real data.</p>
      </div>

      <div class="chat" id="chat"></div>

      <div class="chiprow">
        <button class="chip" data-q="how many students">📊 Stats</button>
        <button class="chip" data-q="top students">🏆 Top XP</button>
        <button class="chip" data-q="who is present today">✅ Present</button>
        <button class="chip" data-q="who is absent">🚫 Absent</button>
        <button class="chip" data-q="house standings">🏠 Houses</button>
        <button class="chip" data-q="students by grade">🎓 Grades</button>
        <button class="chip" data-q="help">💡 Help</button>
      </div>

      <div class="ask">
        <input id="botInput" placeholder="Ask me anything…" autocomplete="off">
        <button id="botSend" title="Send">➤</button>
      </div>
    </section>

    <!-- MAIL -->
    <section class="page" data-page="mail">
      <div class="ptitle">✉ Messages</div>
      <div class="psub">Email a parent directly from the roster</div>

      <div class="field">
        <label>Student</label>
        <select id="mailStudent"><option value="">Select a student…</option></select>
      </div>
      <div class="field">
        <label>Parent email</label>
        <input id="mailTo" placeholder="parent@example.com">
      </div>
      <div class="field">
        <label>Subject</label>
        <input id="mailSubject" value="Message from School">
      </div>
      <div class="field">
        <label>Message</label>
        <textarea id="mailBody" rows="7" placeholder="Write your message…"></textarea>
      </div>
      <div class="row">
        <button class="btn" id="mailTemplate">📄 Template</button>
        <button class="btn primary" id="mailSend" style="flex:1">Send email</button>
      </div>
      <p style="font-size:11.5px;color:var(--faint);margin-top:11px;line-height:1.55">
        Set your SMTP details in Settings first. Nothing is sent until you press Send.
      </p>
    </section>

    <!-- NOTES -->
    <section class="page" data-page="notes">
      <div class="wide-head">
        <div><div class="ptitle">📝 Notepad</div>
          <div class="psub">Notes and whiteboard, kept together like FaceAttend 3.2</div></div>
      </div>
      <div class="pane-switch">
        <span class="pane-label">View</span>
        <button class="pane-btn on" id="notesTab">📝 Notes</button>
        <button class="pane-btn" id="boardTab">🎨 Whiteboard</button>
      </div>

      <div id="notesPane" class="pane-body">
        <div class="notepad-layout">
          <aside class="note-sidebar">
            <button class="btn ok block" id="btnNewNote">＋ New note</button>
            <button class="btn bad block" id="btnDelNote">🗑 Delete note</button>
            <div class="sec">My notes</div>
            <div class="notelist" id="noteList"></div>
          </aside>
          <section class="note-editor">
            <div class="note-editor-top">
              <div contenteditable="true" id="noteTitle">Untitled note</div>
              <button class="btn sm" id="btnSaveNote">💾 Save</button>
            </div>
            <div class="note-editor-tools">
              <button class="tool" type="button" title="Bold">B</button>
              <button class="tool" type="button" title="Italic"><i>I</i></button>
              <span style="color:#9ca3af;font:11px var(--mono);margin-left:4px">plain text · autosaves</span>
            </div>
            <div class="note-page"><textarea id="noteBody" placeholder="Start typing…"></textarea></div>
          </section>
        </div>
      </div>

      <div id="boardPane" class="pane-body" style="display:none">
        <div class="whiteboard-layout">
          <aside class="whiteboard-tools">
            <div class="sec" style="margin-top:0">Colour</div>
            <div class="swatches" id="swatches"></div>
            <div class="field">
              <label>Brush size — <span id="brushVal">4</span>px</label>
              <input type="range" id="brush" min="1" max="40" value="4">
            </div>
            <button class="btn block" id="btnEraser" style="margin-bottom:8px">🩹 Eraser</button>
            <button class="btn bad block" id="btnClear" style="margin-bottom:8px">🗑 Clear board</button>
            <button class="btn ok block" id="btnSaveBoard">💾 Save as PNG</button>
          </aside>
          <div class="whiteboard-stage"><canvas id="board" width="1280" height="720"></canvas></div>
        </div>
      </div>
    </section>

    <!-- AWARDS -->
    <section class="page" data-page="awards">
      <div class="ptitle">🏅 Achievements</div>
      <div class="psub">Badges, certificates, QR passes and the attendance calendar</div>

      <div class="sec" style="margin-top:0">Badges</div>
      <div class="card" style="padding:14px">
        <div class="stat"><span>Badges available</span><b id="bdgTotal">6</b></div>
        <div class="stat"><span>Awarded so far</span><b id="bdgAwarded">0</b></div>
        <div class="stat"><span>Students with badges</span><b id="bdgStudents">0</b></div>
      </div>
      <button class="btn block" id="btnBadges" style="margin-top:8px">🏅 View all badges</button>

      <div class="sec">Certificates</div>
      <div class="field">
        <label>Minimum attendance %</label>
        <input id="certPct" type="number" value="90" min="0" max="100">
      </div>
      <button class="btn ok block" id="btnCerts">📜 Generate certificates</button>
      <p style="font-size:11.5px;color:var(--faint);margin-top:8px;line-height:1.55">
        One landscape PDF per eligible student, saved to the Certificates folder.
      </p>

      <div class="sec">Attendance calendar</div>
      <button class="btn block" id="btnCalendar">📅 Open calendar</button>

      <div class="sec">QR passes</div>
      <div class="field">
        <label>Student</label>
        <select id="qrStudent"><option value="">Select a student…</option></select>
      </div>
      <button class="btn block" id="btnQr">🔳 Show QR code</button>

      <div class="sec">Attendance concerns</div>
      <div class="field">
        <label>Alert below this percentage</label>
        <input id="alertPct" type="number" value="85" min="1" max="100">
      </div>
      <button class="btn warn block" id="btnAlerts">⚠ Who is falling behind</button>

      <div class="sec">Reports</div>
      <button class="btn ok block" id="btnReport">📈 Attendance report</button>
      <p style="font-size:11.5px;color:var(--faint);margin-top:8px;line-height:1.55">
        Per-student percentages over the whole term, with absence reasons.
      </p>

      <div class="sec">Audit log</div>
      <button class="btn block" id="btnAudit">🧾 View audit log</button>
      <button class="btn block" id="btnVerify" style="margin-top:8px">🔐 Verify log integrity</button>
    </section>

    <!-- CONSOLE -->
    <section class="page" data-page="console">
      <div class="wide-head">
        <div><div class="ptitle">📋 Console</div>
          <div class="psub">Interactive FaceAttend 3.2 diagnostics and commands</div></div>
        <button class="btn bad sm" id="btnClearLog">Clear console</button>
      </div>
      <div class="console-layout">
        <div class="terminal-shell">
          <div class="terminal-bar">
            <span class="terminal-lights"><i></i><i></i><i></i></span>
            FACEATTEND / DEBUG CONSOLE <span class="terminal-live">LIVE SESSION</span>
          </div>
          <div class="console" id="consoleLog" role="log" aria-live="polite"></div>
          <div class="console-prompt">
            <span class="prompt">fa@local:~$</span>
            <input id="cmdInput" placeholder="Type /? for commands…" autocomplete="off" spellcheck="false">
            <button id="runCmd" type="button">Run</button>
          </div>
        </div>
        <aside class="console-inspector">
          <div class="card">
            <div class="sec" style="margin-top:0">Session</div>
            <div class="stat"><span>Scans today</span><b id="stScans">0</b></div>
            <div class="stat"><span>Unknown faces</span><b id="stUnknown">0</b></div>
            <div class="stat"><span>Uptime</span><b id="stUptime">0m</b></div>
          </div>
          <div class="card">
            <div class="sec" style="margin-top:0">Quick commands</div>
            <div class="command-list">
              <div><b>/?</b> show help</div>
              <div><b>/list</b> list students</div>
              <div><b>/stats</b> attendance stats</div>
              <div><b>/scan start</b> start camera</div>
              <div><b>/scan stop</b> stop camera</div>
              <div><b>/launch snake</b> open arcade</div>
            </div>
          </div>
        </aside>
      </div>
    </section>

    <!-- SETTINGS -->
    <section class="page" data-page="settings">
      <div class="ptitle">⚙ Settings</div>
      <div class="psub">Customise how FaceAttend behaves</div>

      <div class="sec" style="margin-top:0">Appearance</div>
      <div class="tg">
        <div><div class="tl">Light theme</div><div class="td">Switch to a bright palette</div></div>
        <div class="sw" id="swTheme"></div>
      </div>
      <div class="tg">
        <div><div class="tl">Reduce motion</div><div class="td">Turn off background animation</div></div>
        <div class="sw" id="swMotion"></div>
      </div>
      <div class="tg">
        <div><div class="tl">Sound effects</div><div class="td">Chime on recognition</div></div>
        <div class="sw on" id="swSound"></div>
      </div>
      <div class="tg">
        <div><div class="tl">Confetti</div><div class="td">Celebrate every scan</div></div>
        <div class="sw on" id="swConfetti"></div>
      </div>

      <div class="sec">Attendance rules</div>
      <div class="field">
        <label>Late cutoff (HH:MM)</label>
        <input id="setLate" placeholder="08:30">
      </div>
      <div class="field">
        <label>School name</label>
        <input id="setSchool" placeholder="Your school">
      </div>
      <div class="field">
        <label>Face match threshold — <span id="thVal">0.45</span></label>
        <input type="range" id="setThreshold" min="0.3" max="0.7" step="0.01" value="0.45">
      </div>
      <div class="grid2">
        <div class="field"><label>Confirm over N frames</label>
          <input id="setConfirm" type="number" value="2" min="1" max="6"></div>
        <div class="field"><label>Match margin (0 = off)</label>
          <input id="setMargin" type="number" value="0" min="0" max="0.2" step="0.01"></div>
      </div>
      <p style="font-size:11px;color:var(--faint);margin:-6px 0 12px;line-height:1.55">
        Confirming over two frames removes one-off misreads. Raise the margin
        only if two students are actually being confused.
      </p>
      <div class="grid2">
        <div class="field"><label>XP on time</label><input id="setXpOn" type="number" value="5"></div>
        <div class="field"><label>XP late</label><input id="setXpLate" type="number" value="2"></div>
      </div>

      <div class="sec">Enrolment</div>
      <button class="btn block" id="btnBulkPhotos">🖼 Import photos in bulk</button>
      <p style="font-size:11.5px;color:var(--faint);margin-top:8px;line-height:1.55">
        Name each file after the roll number or UID, e.g. <b>8C01.jpg</b>.
      </p>

      <div class="sec">Parent portal</div>
      <button class="btn block" id="btnPortal">🔗 Open the parent page</button>
      <p style="font-size:11.5px;color:var(--faint);margin-top:8px;line-height:1.55">
        Parents enter their child's code and see only that child.
      </p>

      <div class="sec">Records</div>
      <div class="tg">
        <div><div class="tl">Keep a photo of each scan</div>
          <div class="td">Evidence for the audit log, pruned automatically</div></div>
        <div class="sw on" id="swEvidence"></div>
      </div>
      <div class="tg">
        <div><div class="tl">High contrast</div>
          <div class="td">Stronger colours and heavier text</div></div>
        <div class="sw" id="swContrast"></div>
      </div>

      <div class="sec">Exhibition day</div>
      <div id="demoBanner" style="display:none;padding:11px 13px;border-radius:12px;
           background:rgba(165,107,255,.14);border:1px solid var(--accent2);
           font-size:11.5px;line-height:1.55;margin-bottom:10px">
        <b style="color:var(--accent2)">Demo mode is on</b>
        <div id="demoInfo" style="color:var(--muted);margin-top:4px"></div>
      </div>
      <button class="btn block" id="btnDemoStart">🎪 Start demo mode</button>
      <button class="btn ok block" id="btnDemoRestore" style="margin-top:8px">
        ↩ Restore the real roster
      </button>
      <button class="btn block" id="btnClearToday" style="margin-top:8px">
        🧹 Clear today's scans only
      </button>
      <p style="font-size:11.5px;color:var(--faint);margin-top:8px;line-height:1.55">
        Snapshots the roster, faces and scans. Let visitors enrol and scan
        freely, then restore and nothing they did remains.
      </p>

      <div class="sec">Backups</div>
      <div class="tg">
        <div><div class="tl">Daily automatic backup</div>
          <div class="td">Runs once a day, keeps the most recent copies</div></div>
        <div class="sw on" id="swAutoBackup"></div>
      </div>
      <div class="card" style="padding:13px;margin-top:8px">
        <div class="stat"><span>Last backup</span><b id="bkLast">—</b></div>
        <div class="stat"><span>Copies kept</span><b id="bkKept">—</b></div>
      </div>
      <button class="btn block" id="btnBackupNow" style="margin-top:8px">💾 Back up now</button>

      <div class="sec" id="staffSec">Staff accounts</div>
      <button class="btn block" id="btnStaff">👥 Manage teachers</button>
      <p style="font-size:11.5px;color:var(--faint);margin-top:8px;line-height:1.55">
        Teachers see only their own classes. Admins see everything.
      </p>

      <div class="sec">School timetable</div>
      <div class="tg">
        <div><div class="tl">Enforce timetable</div>
          <div class="td">Use school hours to decide present, late, half day or absent</div></div>
        <div class="sw" id="swEnforce"></div>
      </div>
      <button class="btn block" id="btnSchedule" style="margin-top:8px">
        🕘 Edit timetable &amp; holidays
      </button>
      <button class="btn block" id="btnCloseRegister" style="margin-top:8px">
        📕 Close register now
      </button>
      <p style="font-size:11.5px;color:var(--faint);margin-top:8px;line-height:1.55">
        Closing writes an absence for every student who did not scan today.
      </p>

      <div class="sec">Security &amp; hardware</div>
      <div class="tg">
        <div><div class="tl">Liveness check</div><div class="td">Reject photos held to the camera</div></div>
        <div class="sw on" id="swLiveness"></div>
      </div>
      <div class="tg">
        <div><div class="tl">Voice feedback</div><div class="td">Speak each student's name aloud</div></div>
        <div class="sw on" id="swVoice"></div>
      </div>
      <div class="tg">
        <div><div class="tl">External API</div><div class="td">Serve read-only data on port 8788</div></div>
        <div class="sw" id="swApi"></div>
      </div>
      <div class="field" style="margin-top:12px">
        <label>Camera index (single webcam)</label>
        <input id="setCamIndex" type="number" value="0" min="0" max="8">
      </div>
      <button class="btn block" id="btnCameras" style="margin-bottom:6px">
        📹 Camera sources / IP cameras
      </button>
      <button class="btn block" id="btnRecogTest" style="margin-bottom:6px">
        🎯 Test recognition accuracy
      </button>
      <button class="btn block" id="btnEncrypt" style="margin-top:6px">🔒 Encrypt roster file</button>

      <div class="sec">Email (SMTP)</div>
      <div class="field"><label>From address</label><input id="setSmtpUser" placeholder="you@gmail.com"></div>
      <div class="field">
        <label>App password <span id="pwSaved" style="color:var(--ok);
               font-weight:600;text-transform:none;letter-spacing:0"></span></label>
        <input id="setSmtpPass" type="password" placeholder="••••••••">
        <div style="font-size:11px;color:var(--faint);margin-top:5px;line-height:1.5">
          Saved on this computer. Leave blank to keep the current one.
        </div>
      </div>
      <div class="grid2">
        <div class="field"><label>Host</label><input id="setSmtpHost" value="smtp.gmail.com"></div>
        <div class="field"><label>Port</label><input id="setSmtpPort" type="number" value="587"></div>
      </div>

      <div class="tg" style="margin-top:10px">
        <div><div class="tl">Auto-notify absent</div><div class="td">Email parents at the time below</div></div>
        <div class="sw" id="swAutoNotify"></div>
      </div>
      <div class="field" style="margin-top:10px">
        <label>Notify time (HH:MM)</label>
        <input id="setNotifyTime" value="10:00">
      </div>

      <button class="btn primary block" id="btnSaveSettings" style="margin-top:14px">Save settings</button>
      <button class="btn block" id="btnBackup" style="margin-top:8px">💾 Backup data</button>

      <div class="sec">Account</div>
      <div class="card" style="padding:14px">
        <div class="stat"><span>Signed in as</span><b id="whoami">—</b></div>
      </div>
      <button class="btn block" id="btnChangePw" style="margin-top:8px">🔑 Change password</button>
      <button class="btn block" id="btnSwitch" style="margin-top:8px">🔄 Switch account</button>
      <button class="btn bad block" id="btnSignOut" style="margin-top:8px">Sign out</button>

      <div class="sec">About</div>
      <p style="font-size:12px;color:var(--faint);line-height:1.7">
         FaceAttend 3.2 — web edition.<br>
         Open Console for live diagnostics and slash commands.<br>
        Data lives in <span id="dataPath" style="color:var(--muted)"></span>
      </p>
    </section>
  </aside>

  <!-- ══ MAIN ══ -->
  <main class="main" id="main">

    <!-- dashboard view -->
    <div id="viewDash">
      <div class="head">
        <div>
          <h1>Teacher Dashboard</h1>
          <div class="date"><span id="today"></span><span class="clock" id="clock"></span></div>
        </div>
        <div class="row">
          <button class="btn ok" id="btnAdd">＋ Add</button>
          <button class="btn warn" id="btnEdit">✎ Edit</button>
          <button class="btn primary" id="btnMark">✓ Mark attendance</button>
          <button class="btn bad" id="btnRemove">🗑 Remove</button>
          <button class="btn" id="btnImport">📥 Import</button>
          <button class="btn" id="btnExport">⬇ Export</button>
          <button class="btn" id="btnBulk">📋 Bulk add</button>
          <button class="btn" id="btnSection">🔀 Move</button>
          <button class="btn" id="btnUnmark">↩ Unmark</button>
          <button class="btn" id="btnAbsence">📝 Absence reason</button>
          <button class="btn" id="btnSlip">🧾 Late slip</button>
          <button class="btn" id="btnNotify">📨 Notify absent</button>
        </div>
      </div>

      <div class="kpis" id="kpis"></div>

      <div class="bar">
        <div class="search">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round">
            <circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>
          </svg>
          <input id="q" placeholder="Search students by name, roll, UID, house…" autocomplete="off">
          <kbd>/</kbd>
        </div>
        <select id="fHouse" class="btn" style="padding-right:30px">
          <option value="">All houses</option>
          <option>Red</option><option>Blue</option><option>Green</option><option>Orange</option>
        </select>
        <select id="fStatus" class="btn" style="padding-right:30px">
          <option value="">All students</option>
          <option value="present">Present today</option>
          <option value="absent">Absent today</option>
        </select>
        <button class="btn sm" id="btnSelAll">☑ Select all</button>
        <button class="btn sm ghost" id="btnClearSel">Clear</button>
      </div>

      <div class="tablewrap">
        <div class="scroll">
          <table>
            <thead><tr id="thead"></tr></thead>
            <tbody id="tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

  </main>
</div>

<!-- ══ MODAL ══ -->
<div class="veil" id="veil">
  <div class="modal" id="modal"></div>
</div>

<!-- ══ ARCADE ══ -->
<div class="arcade" id="arcadeVeil">
  <div class="arcadebox" id="arcadeBox"></div>
</div>

<div class="toasts" id="toasts"></div>
<div class="conf" id="conf"></div>

<script src="/static/app.js"></script>
</body>
</html>
"""

PORTAL_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Attendance — FaceAttend</title>
<link rel="stylesheet" href="/static/app.css">
<style>
  body{overflow:auto}
  .wrap{max-width:620px;margin:0 auto;padding:28px 20px 60px}
  .plogo{display:flex;align-items:center;gap:12px;margin-bottom:24px}
  .plogo .m{width:44px;height:44px;border-radius:14px;display:grid;
    place-items:center;background:var(--grad);color:#fff;font-weight:800}
  .pcard{background:var(--surface);border:1px solid var(--stroke);
    border-radius:var(--r-lg);padding:20px;margin-bottom:14px}
  .big{font-size:44px;font-weight:850;letter-spacing:-1.6px;line-height:1}
  .prow{display:flex;justify-content:space-between;padding:10px 0;
    border-bottom:1px solid var(--stroke);font-size:13.5px}
  .prow:last-child{border:0}
  .bar{height:9px;border-radius:5px;background:var(--surface-2);
    overflow:hidden;margin-top:12px}
  .bar i{display:block;height:100%;border-radius:5px;background:var(--grad)}
  .bdg{display:inline-flex;align-items:center;gap:6px;padding:6px 11px;
    border-radius:99px;background:rgba(255,209,102,.13);color:var(--xp);
    font-size:12px;font-weight:650;margin:0 6px 6px 0}
  @media(max-width:520px){ .big{font-size:36px} }
</style>
</head>
<body>
<div class="aurora"><b></b><b></b></div>
<div class="wrap">
  <div class="plogo">
    <div class="m">FA</div>
    <div>
      <div style="font-size:19px;font-weight:800;letter-spacing:-.4px">Attendance</div>
      <div style="font-size:12px;color:var(--muted)" id="pSchool">for parents</div>
    </div>
  </div>

  <div class="pcard" id="pLookup">
    <div class="field" style="margin:0">
      <label>Student code</label>
      <input id="pUid" placeholder="F001" autocomplete="off"
             style="text-transform:uppercase">
    </div>
    <button class="btn primary block" id="pGo" style="margin-top:12px">View attendance</button>
    <p style="font-size:11.5px;color:var(--faint);margin-top:10px;line-height:1.55">
      The code is on your child's QR pass. This page shows only that child.
    </p>
    <div id="pErr" style="color:var(--bad);font-size:12.5px;margin-top:10px"></div>
  </div>

  <div id="pResult" style="display:none"></div>
</div>

<script>
const $$$ = id => document.getElementById(id);
async function look() {
  const uid = $$$('pUid').value.trim().toUpperCase();
  if (!uid) return;
  $$$('pErr').textContent = '';
  try {
    const r = await fetch('/api/portal/' + encodeURIComponent(uid));
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || 'Not found');
    $$$('pSchool').textContent = d.school;
    const band = d.pct >= 90 ? 'var(--ok)' : d.pct >= 75 ? 'var(--warn)' : 'var(--bad)';
    $$$('pResult').style.display = '';
    $$$('pResult').innerHTML = `
      <div class="pcard">
        <div style="font-size:12px;color:var(--muted);text-transform:uppercase;
             letter-spacing:.9px;font-weight:750">Attendance</div>
        <div class="big" style="color:${band}">${d.pct}%</div>
        <div style="font-size:13px;color:var(--muted);margin-top:6px">
          ${d.days_present} of ${d.school_days} school days</div>
        <div class="bar"><i style="width:${Math.min(100, d.pct)}%"></i></div>
      </div>

      <div class="pcard">
        <div class="prow"><span>Student</span><b>${d.name}</b></div>
        <div class="prow"><span>Class</span><b>${d.grade || '—'}</b></div>
        <div class="prow"><span>House</span><b>${d.house || '—'}</b></div>
        <div class="prow"><span>Today</span>
          <b style="color:${d.present_today ? 'var(--ok)' : 'var(--muted)'}">
            ${d.present_today ? 'Present' : 'Not marked in'}</b></div>
        <div class="prow"><span>Late arrivals</span><b>${d.days_late}</b></div>
        <div class="prow"><span>Reward points</span><b>${d.xp} XP</b></div>
      </div>

      ${d.badges.length ? `<div class="pcard">
        <div style="font-size:12px;color:var(--muted);text-transform:uppercase;
             letter-spacing:.9px;font-weight:750;margin-bottom:12px">Badges earned</div>
        ${d.badges.map(b => `<span class="bdg">${b.icon} ${b.name}</span>`).join('')}
      </div>` : ''}

      ${d.absences.length ? `<div class="pcard">
        <div style="font-size:12px;color:var(--muted);text-transform:uppercase;
             letter-spacing:.9px;font-weight:750;margin-bottom:6px">Recorded absences</div>
        ${d.absences.map(a => `<div class="prow"><span>${a.date}</span>
          <b style="font-weight:600;color:var(--muted)">${a.reason}</b></div>`).join('')}
      </div>` : ''}`;
  } catch (e) {
    $$$('pResult').style.display = 'none';
    $$$('pErr').textContent = e.message;
  }
}
$$$('pGo').onclick = look;
$$$('pUid').addEventListener('keydown', e => { if (e.key === 'Enter') look(); });
const q = new URLSearchParams(location.search).get('uid');
if (q) { $$$('pUid').value = q; look(); }
</script>
</body>
</html>
"""

ICON_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128" width="128" height="128">
<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
<stop offset="0%" stop-color="#5b8cff"/><stop offset="55%" stop-color="#a56bff"/>
<stop offset="100%" stop-color="#22d3ee"/></linearGradient></defs>
<rect width="128" height="128" rx="28" fill="#05070d"/>
<rect x="9" y="9" width="110" height="110" rx="23" fill="none" stroke="url(#g)" stroke-width="2.5" opacity=".5"/>
<g fill="none" stroke="url(#g)" stroke-width="6.5" stroke-linecap="round">
<path d="M32 46V37a6 6 0 0 1 6-6h10"/><path d="M96 46V37a6 6 0 0 0-6-6H80"/>
<path d="M32 82v9a6 6 0 0 0 6 6h10"/><path d="M96 82v9a6 6 0 0 1-6 6H80"/></g>
<circle cx="52" cy="57" r="5" fill="#5b8cff"/><circle cx="76" cy="57" r="5" fill="#a56bff"/>
<path d="M50 75c4.5 5.5 9.5 8 14 8s9.5-2.5 14-8" fill="none" stroke="#2ee6a8"
stroke-width="6.5" stroke-linecap="round"/></svg>"""

# ══════════════════════════════════════════════════════════════════════════
#  AI BUDDY  — offline brain, Ollama optional
# ══════════════════════════════════════════════════════════════════════════
import difflib
import urllib.error
import urllib.request


class Buddy:
    def __init__(self, store: Store, scanner: Scanner, settings: Settings):
        self.store = store
        self.scanner = scanner
        self.settings = settings

    # ── helpers ───────────────────────────────────────────────────
    def ollama_up(self) -> bool:
        try:
            req = urllib.request.Request(f"{OLLAMA_URL.rstrip('/')}/api/tags")
            with urllib.request.urlopen(req, timeout=1.5):
                return True
        except Exception:
            return False

    def find(self, text: str, students: list[dict]):
        """Fuzzy student lookup — tolerates typos and partial names."""
        t = (text or "").lower().strip()
        if not t or not students:
            return None
        for s in students:
            for f in ("UID", "Roll No"):
                v = str(s.get(f, "")).lower().strip()
                if v and v in t.split():
                    return s
        best, score = None, 0.0
        for s in students:
            name = str(s.get("Name", "")).lower().strip()
            if not name:
                continue
            if name in t:
                return s
            for part in name.split():
                if len(part) > 2 and part in t:
                    return s
            for w in t.split():
                if len(w) < 3:
                    continue
                r = difflib.SequenceMatcher(None, w, name.split()[0]).ratio()
                if r > score:
                    best, score = s, r
        return best if score >= 0.74 else None

    # ── main entry ────────────────────────────────────────────────
    def answer(self, q: str) -> dict:
        ql = (q or "").lower().strip()
        students = self.store.load_students()
        present = self.scanner.scanned
        pres_names = [s["Name"] for s in students if str(s["UID"]) in present]
        abs_names = [s["Name"] for s in students if str(s["UID"]) not in present]
        R = lambda a, **kw: {"answer": a, "mode": "offline", **kw}

        if not ql:
            return R("Ask me something about the roster.")

        # greetings
        if ql in ("hi", "hello", "hey", "yo", "sup") or ql.startswith(("hi ", "hey ", "hello ")):
            return R("Hey! I can look things up or drive the app for you.\n"
                     'Try "how many students", "who is absent", "search <name>", or "start scan".')
        if "thank" in ql:
            return R("Anytime. Ask me anything else.")
        if any(w in ql for w in ("who are you", "what are you", "your name")):
            return R(f"I'm the AI Buddy inside FaceAttend {APP_VERSION}. I read your live "
                     "roster and can run most of the app for you.")

        # date / time
        if any(w in ql for w in ("what day", "the date", "which day", "today's date")):
            return R(f"Today is {datetime.now():%A, %d %B %Y}.")
        if "time" in ql and any(w in ql for w in ("what", "now", "current")):
            return R(f"It's {datetime.now():%I:%M %p}.")

        # stats
        if any(w in ql for w in ("how many", "total students", "count", "stats", "statistics", "summary")):
            top = max(students, key=lambda s: safe_int(s.get("XP")), default=None)
            return R(f"Total students: {len(students)}\n"
                     f"Present today: {len(pres_names)}\n"
                     f"Absent today: {len(abs_names)}\n"
                     f"XP leader: {top['Name']} ({safe_int(top.get('XP'))} XP)" if top
                     else f"Total students: {len(students)}")

        # leaderboard
        if any(w in ql for w in ("top", "leaderboard", "best", "highest xp", "most xp", "rank")):
            top = sorted(students, key=lambda s: safe_int(s.get("XP")), reverse=True)[:5]
            if not top:
                return R("No students enrolled yet.")
            return R("Top 5 by XP:\n" + "\n".join(
                f"{i+1}. {s['Name']} — {safe_int(s.get('XP'))} XP" for i, s in enumerate(top)))

        # present / absent
        if "absent" in ql:
            if not abs_names:
                return R("Everyone is present. Nice.")
            return R(f"Absent today ({len(abs_names)}):\n" +
                     "\n".join(f"• {n}" for n in abs_names[:20]))
        if "present" in ql or "attendance" in ql:
            if not pres_names:
                return R("Nobody has scanned in yet today.")
            return R(f"Present today ({len(pres_names)}):\n" +
                     "\n".join(f"• {n}" for n in pres_names[:20]))

        # houses
        if "house" in ql:
            st = self.store.house_stats(present)
            ranked = sorted(st.items(), key=lambda x: -x[1]["pct"])
            return R("House standings:\n" + "\n".join(
                f"{i+1}. {h} — {d['pct']:.1f}% ({d['present']}/{d['students']})"
                for i, (h, d) in enumerate(ranked)), navigate="houses")

        # grade breakdown
        if any(w in ql for w in ("by grade", "by section", "by class", "grades", "sections")):
            b: dict[str, int] = {}
            for s in students:
                g = str(s.get("Grade") or "?")
                b[g] = b.get(g, 0) + 1
            if not b:
                return R("No students enrolled yet.")
            return R("Students by grade:\n" + "\n".join(f"• {g}: {n}" for g, n in sorted(b.items())))

        # roster
        if any(w in ql for w in ("list students", "all students", "show roster", "everyone", "who is enrolled")):
            if not students:
                return R("No students yet — press Add to enrol someone.")
            return R(f"All {len(students)} students:\n" + "\n".join(
                f"{i+1}. {s['Name']} — {s.get('Grade') or '?'} · {s.get('House')} · "
                f"{safe_int(s.get('XP'))} XP" for i, s in enumerate(students[:25])))

        # search
        if any(w in ql for w in ("search", "find", "look up", "where is")):
            term = re.sub(r"^(search|find|look\s*up|where\s*is)\s*(for)?\s*", "", ql).strip()
            hit = self.find(term or ql, students)
            if hit:
                return R(self._profile(hit, present), refresh=False)
            return R(f"No student matches '{term}'. Try their full name or UID.")

        # profile
        if any(w in ql for w in ("tell me about", "info on", "details of", "profile", "how is", "days present")):
            hit = self.find(ql, students)
            if hit:
                return R(self._profile(hit, present))
            return R("Which student? Give me a name, roll number or UID.")

        # actions
        if any(w in ql for w in ("start scan", "start camera", "begin scan")):
            ok, err = self.scanner.start()
            return R("Scanner started — watching for faces." if ok else f"Couldn't start: {err}")
        if any(w in ql for w in ("stop scan", "stop camera")):
            self.scanner.stop()
            return R("Scanner stopped.")
        if "export" in ql:
            return R("Use the Export button — I've opened the dashboard for you.", navigate="home")
        for kw, page, label in (("analytic", "analytics", "Analytics"),
                                ("chart", "analytics", "Analytics"),
                                ("whiteboard", "board", "Whiteboard"),
                                ("draw", "board", "Whiteboard"),
                                ("notepad", "notes", "Notepad"),
                                ("note", "notes", "Notepad"),
                                ("setting", "settings", "Settings"),
                                ("email", "mail", "Messages"),
                                ("mail", "mail", "Messages")):
            if kw in ql:
                return R(f"{label} is open.", navigate=page)

        # remove
        if any(w in ql for w in ("remove", "delete")):
            hit = self.find(ql, students)
            if hit:
                return R(f"Found {hit['Name']} ({hit['UID']}). Select their row and press "
                         "Remove — I keep deletion a two-step action on purpose.")
            return R("Which student should I remove? Give me a name or UID.")

        # help
        if any(w in ql for w in ("help", "what can", "commands")):
            return R("Things I'm good at:\n"
                     "• how many students · stats\n"
                     "• top students · who is present · who is absent\n"
                     "• search <name> · tell me about <name>\n"
                     "• house standings · students by grade\n"
                     "• start scan · stop scan\n"
                     "• open analytics / whiteboard / notepad / settings")

        # bare name lookup
        if len(ql.split()) <= 3:
            hit = self.find(ql, students)
            if hit:
                return R(self._profile(hit, present))

        # optional Ollama
        if self.ollama_up():
            txt = self._ollama(q, students, pres_names, abs_names)
            if txt:
                return {"answer": txt, "mode": "AI mode"}

        return R("I didn't catch that. Try:\n"
                 "• how many students\n• who is absent\n• tell me about <name>\n"
                 "• house standings\n• start scan\n\n"
                 "Want free-form answers too? Install Ollama from ollama.com then run:\n"
                 f"  ollama pull {OLLAMA_MODEL}")

    def _profile(self, s: dict, present: set) -> str:
        uid = str(s.get("UID"))
        status = "present today ✓" if uid in present else "not scanned in yet today"
        return (f"{s.get('Name')}\n"
                f"• UID {uid} · Roll {s.get('Roll No') or '?'}\n"
                f"• Grade {s.get('Grade') or '?'} · House {s.get('House')}\n"
                f"• {safe_int(s.get('No. of Days Present'))} days present\n"
                f"• {safe_int(s.get('XP'))} XP · streak {safe_int(s.get('Streak'))}\n"
                f"• Currently {status}")

    def _ollama(self, q, students, pres, absent) -> str:
        roster = "\n".join(
            f"  - {s['Name']} | UID:{s['UID']} | Days:{safe_int(s.get('No. of Days Present'))}"
            f" | XP:{safe_int(s.get('XP'))}" for s in students[:30])
        prompt = (
            f"School: {self.settings['school_name'] or 'School'}\n"
            f"Total students: {len(students)}\n"
            f"Present today ({len(pres)}): {', '.join(pres) or 'none'}\n"
            f"Absent today ({len(absent)}): {', '.join(absent) or 'none'}\n"
            f"Roster:\n{roster}\n\n"
            "You are FaceAttend AI, an assistant for school attendance.\n"
            "Only use data from the roster above. Never invent names or numbers.\n"
            "Answer in 1-3 sentences.\n\n"
            f"Teacher: {q}\nAssistant:")
        try:
            req = urllib.request.Request(
                f"{OLLAMA_URL.rstrip('/')}/api/generate",
                data=json.dumps({"model": OLLAMA_MODEL, "prompt": prompt, "stream": False,
                                 "options": {"temperature": 0.3, "num_predict": 140}}).encode(),
                method="POST", headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=30) as r:
                return str(json.loads(r.read()).get("response", "")).strip()
        except Exception:
            return ""


# ══════════════════════════════════════════════════════════════════════════
#  FLASK APP
# ══════════════════════════════════════════════════════════════════════════
ensure_dirs()
settings = Settings()
store = Store()
badges = None          # set just below, after Scanner exists
scanner = Scanner(store, settings)
badges = Badges(store)
vault = Vault()
buddy = Buddy(store, scanner, settings)

# ══════════════════════════════════════════════════════════════════════════
#  ACCOUNTS
#  Local only — nothing leaves this machine. Passwords are stored as a
#  PBKDF2-HMAC-SHA256 hash with a per-user salt, never in plain text.
# ══════════════════════════════════════════════════════════════════════════
import hashlib
import hmac
import secrets


class Accounts:
    ITERATIONS = 200_000          # deliberately slow, to resist brute force

    def __init__(self):
        self.users: dict = {}
        self.sessions: dict = {}          # token -> {user, created}
        self.load()

    # ── storage ───────────────────────────────────────────────────
    def load(self):
        try:
            if USERS_FILE.exists():
                self.users = json.loads(USERS_FILE.read_text("utf-8"))
        except Exception:
            self.users = {}

    def save(self):
        try:
            USERS_FILE.parent.mkdir(parents=True, exist_ok=True)
            USERS_FILE.write_text(json.dumps(self.users, indent=2), "utf-8")
            try:
                os.chmod(USERS_FILE, 0o600)
            except Exception:
                pass
        except Exception:
            pass

    # ── hashing ───────────────────────────────────────────────────
    def _hash(self, password: str, salt: str) -> str:
        return hashlib.pbkdf2_hmac(
            "sha256", password.encode(), bytes.fromhex(salt), self.ITERATIONS
        ).hex()

    # ── public API ────────────────────────────────────────────────
    def any_users(self) -> bool:
        return bool(self.users)

    def register(self, username: str, password: str, display: str = "") -> tuple[bool, str]:
        username = (username or "").strip().lower()
        if not username or len(username) < 3:
            return False, "Username must be at least 3 characters."
        if username in self.users:
            return False, "That username is already taken."
        if len(password or "") < 6:
            return False, "Password must be at least 6 characters."
        salt = secrets.token_hex(16)
        self.users[username] = {
            "salt": salt,
            "hash": self._hash(password, salt),
            "display": (display or username).strip(),
            "created": datetime.now().isoformat(timespec="seconds"),
            "role": "admin" if not self.users else "teacher",  # first user runs the school
            # Grades this teacher is responsible for. Empty means "all",
            # which is what an admin gets.
            "classes": [],
        }
        self.save()
        return True, ""

    def verify(self, username: str, password: str) -> tuple[bool, str]:
        u = self.users.get((username or "").strip().lower())
        if not u:
            return False, "No account with that username."
        want = u.get("hash", "")
        got = self._hash(password or "", u.get("salt", ""))
        # constant-time compare so timing cannot leak the hash
        if not hmac.compare_digest(want, got):
            return False, "Incorrect password."
        return True, ""


    # ── roles and class scoping ───────────────────────────────────
    def list_users(self) -> list[dict]:
        return [{"username": u, "display": v.get("display", u),
                 "role": v.get("role", "teacher"),
                 "classes": v.get("classes") or [],
                 "created": v.get("created", "")}
                for u, v in sorted(self.users.items())]

    def set_role(self, username: str, role: str) -> tuple[bool, str]:
        username = (username or "").strip().lower()
        if username not in self.users:
            return False, "No such account."
        if role not in ("admin", "teacher"):
            return False, "Role must be admin or teacher."
        # Never leave the school without an administrator.
        if (self.users[username].get("role") == "admin" and role != "admin"
                and sum(1 for v in self.users.values()
                        if v.get("role") == "admin") <= 1):
            return False, "This is the only admin — promote someone else first."
        self.users[username]["role"] = role
        self.save()
        return True, ""

    def set_classes(self, username: str, classes: list) -> tuple[bool, str]:
        username = (username or "").strip().lower()
        if username not in self.users:
            return False, "No such account."
        clean = sorted({str(c).strip() for c in (classes or []) if str(c).strip()})
        self.users[username]["classes"] = clean[:40]
        self.save()
        return True, ""

    def delete_user(self, username: str) -> tuple[bool, str]:
        username = (username or "").strip().lower()
        if username not in self.users:
            return False, "No such account."
        if (self.users[username].get("role") == "admin"
                and sum(1 for v in self.users.values()
                        if v.get("role") == "admin") <= 1):
            return False, "Cannot delete the only admin."
        self.users.pop(username)
        for tok, rec in list(self.sessions.items()):
            if rec.get("user") == username:
                self.sessions.pop(tok, None)
        self.save()
        return True, ""

    def reset_password(self, username: str, new: str) -> tuple[bool, str]:
        username = (username or "").strip().lower()
        if username not in self.users:
            return False, "No such account."
        if len(new or "") < 6:
            return False, "Password must be at least 6 characters."
        u = self.users[username]
        u["salt"] = secrets.token_hex(16)
        u["hash"] = self._hash(new, u["salt"])
        self.save()
        return True, ""

    def start_session(self, username: str) -> str:
        token = secrets.token_urlsafe(32)
        self.sessions[token] = {"user": (username or "").strip().lower(),
                                "created": time.time()}
        return token

    def session_user(self, token: str):
        rec = self.sessions.get(token or "")
        if not rec:
            return None
        # sessions last 12 hours
        if time.time() - rec["created"] > 12 * 3600:
            self.sessions.pop(token, None)
            return None
        return self.users.get(rec["user"]) and dict(
            self.users[rec["user"]], username=rec["user"])

    def end_session(self, token: str):
        self.sessions.pop(token or "", None)

    def change_password(self, username: str, old: str, new: str) -> tuple[bool, str]:
        ok, err = self.verify(username, old)
        if not ok:
            return False, err
        if len(new or "") < 6:
            return False, "New password must be at least 6 characters."
        u = self.users[username.strip().lower()]
        u["salt"] = secrets.token_hex(16)
        u["hash"] = self._hash(new, u["salt"])
        self.save()
        return True, ""


accounts = Accounts()


# ══════════════════════════════════════════════════════════════════════════
#  EXTERNAL HTTP API                                   (ported from 3.2)
#  Read-only endpoints on a second port so other software on the network
#  can pull attendance without touching the main UI server.
# ══════════════════════════════════════════════════════════════════════════
class ApiBridge:
    def __init__(self):
        self._srv = None
        self._thread = None

    def running(self) -> bool:
        return self._srv is not None

    def start(self) -> bool:
        if self._srv is not None:
            return True
        try:
            from http.server import BaseHTTPRequestHandler, HTTPServer
        except Exception:
            return False

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass                              # keep the console quiet

            def _json(self, code, payload):
                body = json.dumps(payload).encode()
                self.send_response(code)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self):
                import urllib.parse as _up
                path = _up.urlparse(self.path).path
                try:
                    if path == "/api/students":
                        self._json(200, {"students": store.load_students()})
                    elif path == "/api/present":
                        self._json(200, {"present": sorted(scanner.scanned)})
                    elif path == "/api/houses":
                        self._json(200, store.house_stats(scanner.scanned))
                    elif path == "/api/stats":
                        st = store.load_students()
                        self._json(200, {
                            "total": len(st),
                            "present": len(scanner.scanned),
                            "absent": max(0, len(st) - len(scanner.scanned)),
                            "scanning": scanner.running,
                            "version": APP_VERSION,
                        })
                    elif path in ("/", "/api"):
                        self._json(200, {
                            "app": "FaceAttend", "version": APP_VERSION,
                            "endpoints": ["/api/students", "/api/present",
                                          "/api/houses", "/api/stats"],
                        })
                    else:
                        self._json(404, {"error": "not found"})
                except Exception as exc:
                    self._json(500, {"error": str(exc)})

        try:
            self._srv = HTTPServer(("0.0.0.0", API_PORT), Handler)
            self._thread = threading.Thread(target=self._srv.serve_forever,
                                            daemon=True)
            self._thread.start()
            return True
        except Exception:
            self._srv = None
            return False

    def stop(self):
        if self._srv is not None:
            try:
                self._srv.shutdown()
                self._srv.server_close()
            except Exception:
                pass
            self._srv = None


api_bridge = ApiBridge()


# ══════════════════════════════════════════════════════════════════════════
#  AUTO-NOTIFY SCHEDULER
#  The "auto-notify absent" setting existed but nothing ever acted on it —
#  which is why the automatic email never arrived. This thread watches the
#  clock and sends once per day at the configured time.
# ══════════════════════════════════════════════════════════════════════════
class NotifyScheduler:
    def __init__(self):
        self.last_sent_date = ""
        self.last_result = ""
        self._load()
        self.thread = threading.Thread(target=self._loop, daemon=True)
        self.thread.start()

    def _state_file(self) -> Path:
        return DATA_DIR / "notify_state.json"

    def _load(self):
        try:
            f = self._state_file()
            if f.exists():
                blob = json.loads(f.read_text("utf-8"))
                self.last_sent_date = blob.get("date", "")
                self.last_result = blob.get("result", "")
        except Exception:
            pass

    def _save(self):
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            self._state_file().write_text(json.dumps(
                {"date": self.last_sent_date, "result": self.last_result}), "utf-8")
        except Exception:
            pass

    def due(self) -> bool:
        if not settings["auto_notify_absent"]:
            return False
        today = str(date.today())
        if self.last_sent_date == today:
            return False                      # already sent today
        try:
            hh, mm = str(settings["absent_notify_time"] or "10:00").split(":")
            target = datetime.now().replace(hour=int(hh), minute=int(mm),
                                            second=0, microsecond=0)
        except Exception:
            return False
        # Fire at or after the target, but not more than two hours late —
        # starting the app in the evening should not trigger a morning email.
        now = datetime.now()
        return target <= now <= target + timedelta(hours=2)

    def send_now(self) -> dict:
        """Email every absent student's parent. Shared by the scheduler and
        the manual button, so both behave identically."""
        import smtplib
        from email.message import EmailMessage
        user, pw = settings["smtp_user"], settings["smtp_pass"]
        if not user or not pw:
            return {"ok": False, "error": "SMTP details are not set."}

        students = store.load_students()
        absent = [x for x in students if str(x.get("UID")) not in scanner.scanned]
        sent, skipped = [], []
        school = settings["school_name"] or "School"
        try:
            with smtplib.SMTP(settings["smtp_host"],
                              safe_int(settings["smtp_port"], 587), timeout=25) as srv:
                srv.starttls()
                srv.login(user, pw)
                for st in absent:
                    to = str(st.get("Parent Email") or "").strip()
                    if not to:
                        skipped.append(st.get("Name"))
                        continue
                    m = EmailMessage()
                    m["From"], m["To"] = user, to
                    m["Subject"] = f"Absence notice — {st.get('Name')}"
                    m.set_content(
                        f"Dear Parent,\n\n{st.get('Name')} was not marked present "
                        f"at {school} today ({datetime.now():%d %B %Y}).\n\n"
                        f"If this is unexpected, please contact the school office.\n\n"
                        f"Kind regards,\n{school}")
                    srv.send_message(m)
                    sent.append(st.get("Name"))
        except Exception as e:
            return {"ok": False, "error": str(e)}
        return {"ok": True, "sent": sent, "skipped": skipped}

    def _loop(self):
        while True:
            try:
                if self.due():
                    res = self.send_now()
                    if res.get("ok"):
                        self.last_sent_date = str(date.today())
                        self.last_result = f"{len(res.get('sent', []))} sent"
                        self._save()
                        print(f"  auto-notify: {self.last_result}")
                    else:
                        # Do NOT mark the day as done on failure, or a brief
                        # network glitch means no email at all that day. The
                        # two-hour window in due() bounds the retries.
                        self.last_result = f"retrying — {res.get('error', 'failed')}"
                        print(f"  auto-notify failed, will retry: "
                              f"{res.get('error', '')[:70]}")
                        time.sleep(180)
            except Exception:
                pass
            time.sleep(45)

    def status(self) -> dict:
        return {
            "enabled": bool(settings["auto_notify_absent"]),
            "time": settings["absent_notify_time"],
            "last_sent_date": self.last_sent_date,
            "last_result": self.last_result,
            "configured": bool(settings["smtp_user"] and settings["smtp_pass"]),
        }


notifier = NotifyScheduler()
# ══════════════════════════════════════════════════════════════════════════
#  REGISTER CLOSE
#  Until now "absent" only meant "has not scanned yet", so at 8am the whole
#  school read absent and nothing was ever written down. Once the register
#  closes, absence becomes a recorded fact with a timestamp.
# ══════════════════════════════════════════════════════════════════════════
class RegisterCloser:
    def __init__(self):
        self.closed_date = ""
        self._load()
        threading.Thread(target=self._loop, daemon=True).start()

    def _file(self) -> Path:
        return DATA_DIR / "register_state.json"

    def _load(self):
        try:
            if self._file().exists():
                self.closed_date = json.loads(
                    self._file().read_text("utf-8")).get("date", "")
        except Exception:
            pass

    def _save(self):
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            self._file().write_text(json.dumps({"date": self.closed_date}), "utf-8")
        except Exception:
            pass

    def close_now(self, actor: str = "system") -> dict:
        """Write an absence row for everyone who never scanned today."""
        students = store.load_students()
        absent = [x for x in students if str(x.get("UID")) not in scanner.scanned]
        for st in absent:
            scanner._audit(str(st.get("UID")), st.get("Name", ""), False, 0,
                           action="absent", reason="Register closed", actor=actor)
        self.closed_date = str(date.today())
        self._save()
        return {"ok": True, "absent": [x.get("Name") for x in absent],
                "present": len(scanner.scanned)}

    def _loop(self):
        while True:
            try:
                if (settings["schedule"] or {}).get("enforce"):
                    today = str(date.today())
                    if self.closed_date != today and is_school_day():
                        info = attendance_status()
                        if info["status"] == "absent":      # past closing time
                            res = self.close_now()
                            print(f"  register closed: {len(res['absent'])} "
                                  f"absence(s) recorded")
            except Exception:
                pass
            time.sleep(60)

    def status(self) -> dict:
        return {"closed_today": self.closed_date == str(date.today()),
                "closed_date": self.closed_date}


register = RegisterCloser()

# ══════════════════════════════════════════════════════════════════════════
#  LATE SLIPS
#  Schools already hand these out on paper. Same PDF machinery as the
#  certificates, so this was mostly free.
# ══════════════════════════════════════════════════════════════════════════
def generate_late_slip(path, name, roll, grade, when, status, reason, school) -> bool:
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A5, landscape
        from reportlab.pdfgen import canvas as rl_canvas
    except Exception:
        return False
    W, H = landscape(A5)
    amber = colors.HexColor("#b8860b")
    ink = colors.HexColor("#243b53")
    try:
        c = rl_canvas.Canvas(str(path), pagesize=landscape(A5))
        c.setFillColor(colors.HexColor("#fffdf5"))
        c.rect(0, 0, W, H, fill=1, stroke=0)
        c.setStrokeColor(amber)
        c.setLineWidth(3)
        c.rect(16, 16, W - 32, H - 32, fill=0, stroke=1)

        c.setFillColor(amber)
        c.setFont("Helvetica-Bold", 22)
        c.drawCentredString(W / 2, H - 62, "LATE ARRIVAL SLIP")
        c.setStrokeColor(amber)
        c.setLineWidth(1.2)
        c.line(W / 2 - 110, H - 74, W / 2 + 110, H - 74)

        c.setFillColor(ink)
        c.setFont("Helvetica", 11)
        y = H - 108
        for label, value in (("Student", str(name)),
                             ("Roll number", str(roll or "-")),
                             ("Class", str(grade or "-")),
                             ("Arrived", str(when)),
                             ("Recorded as", str(status)),
                             ("Reason", str(reason or "not given"))):
            c.setFont("Helvetica-Bold", 10)
            c.drawString(46, y, f"{label}:")
            c.setFont("Helvetica", 11)
            c.drawString(150, y, value)
            y -= 22

        c.setFont("Helvetica-Oblique", 9)
        c.drawString(46, 44, "Present this slip to your class teacher.")
        c.setFont("Helvetica", 9)
        c.drawRightString(W - 46, 44, str(school or "School"))
        c.showPage()
        c.save()
        return True
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════════════
#  OFFLINE QUEUE
#  If the app is killed mid-morning, scans held only in memory are lost.
#  Every recognition is written here first and cleared once it is safely in
#  the roster, so a crash costs nothing.
# ══════════════════════════════════════════════════════════════════════════
class ScanQueue:
    def __init__(self):
        self.pending: list[dict] = []
        self.lock = threading.Lock()
        self.load()

    def load(self):
        try:
            if QUEUE_FILE.exists():
                self.pending = json.loads(QUEUE_FILE.read_text("utf-8")) or []
        except Exception:
            self.pending = []

    def _save(self):
        try:
            QUEUE_FILE.parent.mkdir(parents=True, exist_ok=True)
            QUEUE_FILE.write_text(json.dumps(self.pending), "utf-8")
        except Exception:
            pass

    def add(self, item: dict):
        with self.lock:
            self.pending.append(item)
            self._save()

    def done(self, item: dict):
        with self.lock:
            # `at` of None means "clear every entry for this student",
            # which is what the writer thread wants once the row is saved.
            self.pending = [x for x in self.pending
                            if not (x.get("uid") == item.get("uid")
                                    and (item.get("at") is None
                                         or x.get("at") == item.get("at")))]
            self._save()

    def replay(self) -> int:
        """Re-apply anything left over from a previous run."""
        with self.lock:
            left = list(self.pending)
        if not left:
            return 0
        done = 0
        for item in left:
            uid = str(item.get("uid", ""))
            if item.get("date") != str(date.today()):
                self.done(item)          # stale, belongs to another day
                continue
            if uid in scanner.scanned:
                self.done(item)
                continue
            rec = store.award(uid, safe_int(item.get("xp"), 0))
            if rec:
                scanner.scanned.add(uid)
                scanner.save_scans()
                scanner._audit(uid, rec.get("Name", uid),
                               item.get("status") in ("late", "half_day"),
                               safe_int(item.get("xp"), 0),
                               reason=item.get("status", ""), actor="recovered")
                done += 1
            self.done(item)
        if done:
            print(f"  recovered {done} scan(s) from the previous session")
        return done

    def count(self) -> int:
        return len(self.pending)


queue = ScanQueue()


# ══════════════════════════════════════════════════════════════════════════
#  EXHIBITION MODE
#  Visitors will scan their own faces. Without this you would be deleting
#  rows by hand between demonstrations, and one slip could take the real
#  roster with it. Snapshot first, play freely, restore afterwards.
# ══════════════════════════════════════════════════════════════════════════
class DemoMode:
    def __init__(self):
        self.dir = DATA_DIR / "demo_snapshot"

    def active(self) -> bool:
        return (self.dir / "marker.json").exists()

    def info(self) -> dict:
        try:
            if self.active():
                return json.loads((self.dir / "marker.json").read_text("utf-8"))
        except Exception:
            pass
        return {}

    def start(self, actor: str = "") -> dict:
        """Copy the roster, faces, encodings and today's scans aside."""
        try:
            if self.dir.exists():
                shutil.rmtree(self.dir)
            self.dir.mkdir(parents=True, exist_ok=True)
            if EXCEL_FILE.exists():
                shutil.copy2(EXCEL_FILE, self.dir / EXCEL_FILE.name)
            if FACES_DIR.exists():
                shutil.copytree(FACES_DIR, self.dir / "faces", dirs_exist_ok=True)
            if ENCODINGS_FILE.exists():
                shutil.copy2(ENCODINGS_FILE, self.dir / ENCODINGS_FILE.name)
            if SCANS_FILE.exists():
                shutil.copy2(SCANS_FILE, self.dir / SCANS_FILE.name)
            if BADGE_FILE.exists():
                shutil.copy2(BADGE_FILE, self.dir / BADGE_FILE.name)
            students = store.load_students()
            (self.dir / "marker.json").write_text(json.dumps({
                "started": datetime.now().isoformat(timespec="seconds"),
                "students": len(students), "by": actor,
            }), "utf-8")
            return {"ok": True, "students": len(students)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def restore(self) -> dict:
        """Put everything back exactly as it was before the demo."""
        if not self.active():
            return {"ok": False, "error": "No snapshot to restore."}
        try:
            snap_xl = self.dir / EXCEL_FILE.name
            if snap_xl.exists():
                shutil.copy2(snap_xl, EXCEL_FILE)
            if (self.dir / "faces").exists():
                if FACES_DIR.exists():
                    shutil.rmtree(FACES_DIR)
                shutil.copytree(self.dir / "faces", FACES_DIR)
            for f in (ENCODINGS_FILE, SCANS_FILE, BADGE_FILE):
                src = self.dir / f.name
                if src.exists():
                    shutil.copy2(src, f)
                elif f.exists():
                    f.unlink()
            scanner.load_encodings()
            scanner.scanned.clear()
            scanner.load_scans()
            badges.load()
            students = store.load_students()
            shutil.rmtree(self.dir)
            return {"ok": True, "students": len(students)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def clear_today(self) -> dict:
        """Wipe today's scans only, leaving the roster untouched."""
        n = len(scanner.scanned)
        scanner.scanned.clear()
        scanner.save_scans()
        return {"ok": True, "cleared": n}


demo = DemoMode()


# ══════════════════════════════════════════════════════════════════════════
#  AUTOMATIC BACKUPS
#  Backing up was a button nobody remembers to press. One bad moment during
#  a demonstration and the roster is gone, so this runs on its own.
# ══════════════════════════════════════════════════════════════════════════
class BackupJob:
    def __init__(self):
        self.last_date = ""
        self.last_result = ""
        self._load()
        threading.Thread(target=self._loop, daemon=True).start()

    def _file(self) -> Path:
        return DATA_DIR / "backup_state.json"

    def _load(self):
        try:
            if self._file().exists():
                blob = json.loads(self._file().read_text("utf-8"))
                self.last_date = blob.get("date", "")
                self.last_result = blob.get("result", "")
        except Exception:
            pass

    def _save(self):
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            self._file().write_text(json.dumps(
                {"date": self.last_date, "result": self.last_result}), "utf-8")
        except Exception:
            pass

    def run(self) -> dict:
        try:
            BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
            out = BACKUPS_DIR / f"auto_{datetime.now():%Y%m%d_%H%M%S}"
            out.mkdir(exist_ok=True)
            if EXCEL_FILE.exists():
                shutil.copy2(EXCEL_FILE, out / EXCEL_FILE.name)
            if FACES_DIR.exists():
                shutil.copytree(FACES_DIR, out / "faces", dirs_exist_ok=True)
            for f in (ENCODINGS_FILE, SETTINGS_FILE, BADGE_FILE, AUDIT_LOG):
                if f.exists():
                    shutil.copy2(f, out / f.name)
            self._prune()
            return {"ok": True, "path": str(out)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _prune(self):
        """Keep only the most recent N automatic backups."""
        keep = max(2, safe_int(settings["backup_keep"], 14))
        try:
            autos = sorted([d for d in BACKUPS_DIR.iterdir()
                            if d.is_dir() and d.name.startswith("auto_")])
            for old in autos[:-keep]:
                shutil.rmtree(old, ignore_errors=True)
        except Exception:
            pass

    def _loop(self):
        while True:
            try:
                if settings["auto_backup"]:
                    today = str(date.today())
                    if self.last_date != today:
                        res = self.run()
                        self.last_date = today
                        self.last_result = ("saved" if res.get("ok")
                                            else res.get("error", "failed"))
                        self._save()
                        if res.get("ok"):
                            print(f"  daily backup: {res['path']}")
            except Exception:
                pass
            time.sleep(600)

    def status(self) -> dict:
        count = 0
        try:
            count = len([d for d in BACKUPS_DIR.iterdir()
                         if d.is_dir() and d.name.startswith("auto_")])
        except Exception:
            pass
        return {"enabled": bool(settings["auto_backup"]),
                "last_date": self.last_date, "last_result": self.last_result,
                "kept": count, "keep_limit": safe_int(settings["backup_keep"], 14)}


backups = BackupJob()




app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


@app.after_request
def no_cache(r):
    r.headers["Cache-Control"] = "no-store"
    return r


@app.route("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html")


@app.route("/static/app.css")
def css():
    return Response(APP_CSS, mimetype="text/css")


@app.route("/static/app.js")
def js():
    return Response(APP_JS, mimetype="application/javascript")


@app.route("/static/icon.svg")
def icon():
    return Response(ICON_SVG, mimetype="image/svg+xml")


def attendance_trend(total: int, days: int = 30) -> list[float]:
    """Daily attendance % for the last N days, read from the audit log."""
    if not AUDIT_LOG.exists() or total <= 0:
        return []
    seen: dict[str, set] = {}
    try:
        with open(AUDIT_LOG, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                d = row.get("date")
                if d:
                    seen.setdefault(d, set()).add(row.get("uid"))
    except Exception:
        return []
    if not seen:
        return []
    today = date.today()
    out = []
    for i in range(days - 1, -1, -1):
        d = str(today - timedelta(days=i))
        out.append(round(len(seen.get(d, ())) / total * 100, 1))
    # trim leading run of days with no records at all
    first = next((i for i, v in enumerate(out) if v > 0), None)
    return out[max(0, first - 1):] if first is not None else []


def current_user():
    return accounts.session_user(_token())


def is_admin(user=None) -> bool:
    u = user or current_user()
    return bool(u and u.get("role") == "admin")


def visible_classes(user=None) -> list:
    """
    Grades this account may see. An empty list means everything, which is
    what admins and unrestricted teachers get.
    """
    u = user or current_user()
    if not u or u.get("role") == "admin":
        return []
    return list(u.get("classes") or [])


def scope_students(students: list[dict], user=None) -> list[dict]:
    """Filter a roster down to the classes this account is responsible for."""
    allowed = visible_classes(user)
    if not allowed:
        return students
    wanted = {c.strip().lower() for c in allowed}
    return [s for s in students
            if str(s.get("Grade") or "").strip().lower() in wanted]


def may_touch(uid: str, user=None) -> bool:
    """Is this student inside the account's classes?"""
    allowed = visible_classes(user)
    if not allowed:
        return True
    for s in store.load_students():
        if str(s.get("UID")) == str(uid):
            return str(s.get("Grade") or "").strip().lower() in {
                c.strip().lower() for c in allowed}
    return False


def require_admin():
    """Return an error response when the caller is not an admin, else None."""
    if is_admin():
        return None
    return jsonify({"error": "Only an administrator can do that."}), 403


# ── accounts ──────────────────────────────────────────────────────
def _token() -> str:
    auth = request.headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        return auth[7:]
    return request.cookies.get("fa_session", "")


@app.route("/api/auth/profiles")
def api_auth_profiles():
    """
    Names of the accounts on this machine, for the sign-in picker.

    Only the display name and role are exposed — no password material and
    nothing that identifies a student.
    """
    return jsonify({"profiles": [
        {"username": u, "display": v.get("display", u),
         "role": v.get("role", "teacher")}
        for u, v in sorted(accounts.users.items())]})


@app.route("/api/auth/status")
def api_auth_status():
    user = accounts.session_user(_token())
    return jsonify({
        "setup_needed": not accounts.any_users(),
        "signed_in": user is not None,
        "user": {"username": user["username"], "display": user.get("display"),
                 "role": user.get("role")} if user else None,
    })


@app.route("/api/auth/register", methods=["POST"])
def api_auth_register():
    d = request.json or {}
    ok, err = accounts.register(d.get("username", ""), d.get("password", ""),
                                d.get("display", ""))
    if not ok:
        return jsonify({"error": err}), 400
    uname = (d.get("username", "") or "").strip().lower()
    token = accounts.start_session(uname)
    u = accounts.users.get(uname, {})
    resp = jsonify({"ok": True, "token": token, "username": uname,
                    "display": u.get("display", uname),
                    "role": u.get("role", "teacher")})
    resp.set_cookie("fa_session", token, httponly=True, samesite="Lax",
                    max_age=12 * 3600)
    return resp


@app.route("/api/auth/login", methods=["POST"])
def api_auth_login():
    d = request.json or {}
    ok, err = accounts.verify(d.get("username", ""), d.get("password", ""))
    if not ok:
        # Same delay on every failure so timing cannot reveal whether the
        # username exists.
        time.sleep(0.4)
        return jsonify({"error": err}), 401
    uname = (d.get("username", "") or "").strip().lower()
    token = accounts.start_session(uname)
    u = accounts.users.get(uname, {})
    resp = jsonify({"ok": True, "token": token, "username": uname,
                    "display": u.get("display", uname),
                    "role": u.get("role", "teacher")})
    resp.set_cookie("fa_session", token, httponly=True, samesite="Lax",
                    max_age=12 * 3600)
    return resp


@app.route("/api/auth/logout", methods=["POST"])
def api_auth_logout():
    accounts.end_session(_token())
    resp = jsonify({"ok": True})
    resp.set_cookie("fa_session", "", expires=0)
    return resp


@app.route("/api/auth/password", methods=["POST"])
def api_auth_password():
    user = accounts.session_user(_token())
    if not user:
        return jsonify({"error": "Not signed in."}), 401
    d = request.json or {}
    ok, err = accounts.change_password(user["username"], d.get("old", ""),
                                       d.get("new", ""))
    if not ok:
        return jsonify({"error": err}), 400
    return jsonify({"ok": True})


# ── state ─────────────────────────────────────────────────────────
@app.route("/api/state")
def api_state():
    me = current_user()
    all_students = store.load_students()
    students = scope_students(all_students, me)
    return jsonify({
        "students": students,
        "role": (me or {}).get("role", ""),
        "my_classes": visible_classes(me),
        "scoped": len(students) != len(all_students),
        "trend": attendance_trend(len(students)),
        "present": sorted(uid for uid in scanner.scanned
                          if any(str(x.get("UID")) == uid for x in students)),
        "settings": {**settings.data, "smtp_pass": ""},
        "smtp_pass_saved": bool(settings["smtp_pass"]),
        "houses": store.house_stats(scanner.scanned, students),
        "scanning": scanner.running,
        "unknown_count": scanner.unknown_count,
        "data_dir": str(APP_DIR),
        "liveness_rejects": scanner.liveness_rejects,
        "ambiguous": scanner.ambiguous,
        "api_running": api_bridge.running(),
        "api_port": API_PORT,
        "vault_available": vault.available(),
        "badge_defs": Badges.DEFS,
        "missing": MISSING,
        "face_engine": bool(face_recognition is not None),
        "enrolled_faces": len(scanner.encodings),
        "recognition_ready": bool(face_recognition is not None and scanner.encodings),
        "attendance": attendance_status(),
        "queued": queue.count(),
        "evidence_on": bool(settings["keep_scan_photos"]),
        "version": APP_VERSION,
        "sounds": [k for k, p in SOUND_FILES.items() if p.exists()],
    })


@app.route("/api/sound/<key>")
def api_sound(key):
    path = SOUND_FILES.get(key)
    if not path or not path.exists():
        return jsonify({"error": "not found"}), 404
    return send_file(str(path))


@app.route("/api/settings", methods=["POST"])
def api_settings():
    denied = require_admin()
    if denied:
        return denied
    settings.save(request.json or {})
    return jsonify({"ok": True})


# ── students ──────────────────────────────────────────────────────
@app.route("/api/students", methods=["POST"])
def api_add():
    d = request.json or {}
    problems = validate_student(d, store.load_students())
    if problems:
        return jsonify({"error": " ".join(problems)}), 400
    uid = store.add(d)

    photos = [d.get("photo")] + list(d.get("extra_photos") or [])
    photos = [x for x in photos if x and "," in x]
    learned = 0
    for i, ph in enumerate(photos):
        try:
            raw = base64.b64decode(ph.split(",", 1)[1])
            # The first shot is the profile picture; the rest are extra
            # angles used only to improve recognition.
            path = FACES_DIR / (f"{uid}.jpg" if i == 0 else f"{uid}_{i}.jpg")
            path.write_bytes(raw)
            if scanner.learn(uid, path, append=(i > 0)):
                learned += 1
        except Exception:
            continue
    if photos and learned == 0:
        return jsonify({"ok": True, "uid": uid,
                        "warn": "No face was detected in those photos."})
    return jsonify({"ok": True, "uid": uid, "encodings": learned})


@app.route("/api/students/<uid>", methods=["PUT"])
def api_update(uid):
    d = request.json or {}
    problems = validate_student(d, store.load_students(), editing_uid=uid)
    if problems:
        return jsonify({"error": " ".join(problems)}), 400
    if not store.update(uid, d):
        return jsonify({"error": "Student not found"}), 404
    return jsonify({"ok": True})


@app.route("/api/students/<uid>", methods=["DELETE"])
def api_delete(uid):
    denied = require_admin()
    if denied:
        return denied
    if not store.remove(uid):
        return jsonify({"error": "Student not found"}), 404
    scanner.encodings.pop(uid, None)
    scanner.save_encodings()
    scanner.scanned.discard(uid)
    scanner.save_scans()
    return jsonify({"ok": True})


@app.route("/api/photo/<uid>")
def api_photo(uid):
    for ext in ("jpg", "jpeg", "png"):
        p = FACES_DIR / f"{uid}.{ext}"
        if p.exists():
            return send_file(p)
    return ("", 404)


@app.route("/api/mark", methods=["POST"])
def api_mark():
    uid = str((request.json or {}).get("uid", ""))
    if not may_touch(uid):
        return jsonify({"error": "That student is not in your classes."}), 403
    if uid in scanner.scanned:
        # This used to return ok:True, so the UI cheerfully said "marked"
        # and it looked like you could add attendance twice. Say no clearly.
        names = {str(x.get("UID")): x.get("Name") for x in store.load_students()}
        return jsonify({"error": f"{names.get(uid, 'That student')} is already "
                                 "marked present today.", "already": True}), 409
    info = attendance_status()
    if not info["can_mark"] and not (request.json or {}).get("force"):
        return jsonify({
            "error": f"Cannot mark now — {info['label'].lower()} ({info['window']}).",
            "status": info["status"], "needs_force": True,
        }), 409
    late = info["status"] in ("late", "half_day")
    xp = info["xp"]
    rec = store.award(uid, xp)
    if not rec:
        return jsonify({"error": "Student not found"}), 404
    scanner.scanned.add(uid)
    scanner.save_scans()
    scanner._audit(uid, rec.get("Name", uid), late, xp, reason=info["status"])
    # Manual marking earns badges exactly as a camera scan does.
    won = badges.check_and_award(uid, rec, scanner.scanned)
    return jsonify({"ok": True, "xp": xp, "late": late,
                    "status": info["status"], "label": info["label"],
                    "badges": [b["name"] for b in won]})


# ── scanning ──────────────────────────────────────────────────────
@app.route("/api/scan/start", methods=["POST"])
def api_scan_start():
    ok, err = scanner.start()
    return jsonify({"ok": ok, "error": err}), (200 if ok else 500)


@app.route("/api/scan/stop", methods=["POST"])
def api_scan_stop():
    scanner.stop()
    return jsonify({"ok": True})


@app.route("/api/scan/last")
def api_scan_last():
    return jsonify(scanner.last)


@app.route("/api/stream")
def api_stream():
    def gen():
        while scanner.running:
            buf = scanner.jpeg()
            if buf:
                yield b"--f\r\nContent-Type: image/jpeg\r\n\r\n" + buf + b"\r\n"
            time.sleep(0.05)
    return Response(gen(), mimetype="multipart/x-mixed-replace; boundary=f")


@app.route("/api/capture", methods=["POST"])
def api_capture():
    frame = scanner.grab_still()
    if frame is None:
        return jsonify({"ok": False, "error": "No camera available"}), 500
    ok, buf = cv2.imencode(".jpg", frame)
    if not ok:
        return jsonify({"ok": False, "error": "Encode failed"}), 500
    return jsonify({"ok": True,
                    "image": "data:image/jpeg;base64," + base64.b64encode(buf).decode()})


# ── buddy ─────────────────────────────────────────────────────────
@app.route("/api/ollama")
def api_ollama():
    """Lets the UI show the right badge before the first question."""
    return jsonify({"up": buddy.ollama_up(), "model": OLLAMA_MODEL, "url": OLLAMA_URL})


@app.route("/api/ask", methods=["POST"])
def api_ask():
    return jsonify(buddy.answer((request.json or {}).get("q", "")))


# ── interactive 3.2 console ───────────────────────────────────────
def _console_lines(*items):
    return [{"text": text, "kind": kind, "time": datetime.now().strftime("%H:%M:%S")}
            for text, kind in items]


@app.route("/api/command", methods=["POST"])
def api_command():
    """Small, safe command dispatcher for the browser console.

    It intentionally mirrors the useful FaceAttend 3.2 desktop commands
    without exposing a shell or allowing arbitrary process execution.
    """
    raw = str((request.json or {}).get("command", "")).strip()
    cmd = raw.lower()
    if not raw:
        return jsonify({"error": "empty command"}), 400

    lines = []
    result = {"ok": True, "lines": lines}

    def say(text, kind="i"):
        lines.extend(_console_lines((text, kind)))

    if cmd in ("/?", "?", "/help", "help"):
        say("FaceAttend 3.2 command reference", "s")
        say("/list          — list all students")
        say("/add           — open Add Student")
        say("/edit <name>   — open a student for editing")
        say("/remove <name> — remove a student from the dashboard")
        say("/refresh       — refresh dashboard data")
        say("/scan start    — start camera scan")
        say("/scan stop     — stop camera scan")
        say("/stats         — show quick stats")
        say("/backup        — backup Excel, faces and settings")
        say("/export        — export the roster")
        say("/badge <name>  — show a student's badges")
        say("/certs         — generate attendance certificates")
        say("/qr <name>     — show a student's QR code")
        say("/calendar      — open the attendance calendar")
        say("/audit         — show the recent audit log")
        say("/notifyabsent  — email every absent student's parent")
        say("/api           — toggle the external HTTP API")
        say("/theme         — toggle light / dark")
        say("/restart       — reload the interface")
        say("/clear         — clear this console")
    elif cmd in ("/list", "list"):
        students = store.load_students()
        say(f"{len(students)} student(s):", "s")
        for s in students:
            say(f"  {s.get('UID','')} | {s.get('Name','')} | "
                f"{s.get('Grade','')} | {s.get('No. of Days Present',0)} days")
    elif cmd in ("/stats", "stats"):
        students = store.load_students()
        present = len([u for u in scanner.scanned
                       if any(str(s.get("UID")) == u for s in students)])
        top = max(students, key=lambda s: safe_int(s.get("XP")), default=None)
        say(f"Total: {len(students)} | Present: {present} | "
            f"Absent: {max(0, len(students) - present)}", "s")
        if top:
            say(f"XP leader: {top.get('Name','')} ({safe_int(top.get('XP'))} XP)")
    elif cmd in ("/refresh", "refresh"):
        result["refresh"] = True
        say("Dashboard state refreshed.", "s")
    elif cmd in ("/scan start", "scan start"):
        ok, err = scanner.start()
        if ok:
            result["refresh"] = True
            say("Scanner started — watching for faces.", "s")
        else:
            say(f"Scanner failed: {err}", "e")
    elif cmd in ("/scan stop", "scan stop"):
        scanner.stop()
        result["refresh"] = True
        say("Scanner stopped.", "w")
    elif cmd.startswith(("/badge", "badge ")):
        term = raw.split(" ", 1)[1].strip().lower() if " " in raw else ""
        students = store.load_students()
        hit = None
        for st in students:
            if term and term in str(st.get("Name", "")).lower():
                hit = st
                break
        if not hit:
            say("Usage: /badge <student name>", "w")
        else:
            got = badges.for_student(str(hit.get("UID")))
            say(f"{hit.get('Name')} — {len(got)}/{len(Badges.DEFS)} badges", "s")
            for b in badges.all_with_progress(hit):
                mark = "\u2713" if b["earned"] else "\u00b7"
                say(f"  {mark} {b['icon']} {b['name']} — {b['detail']}",
                    "s" if b["earned"] else "i")
    elif cmd in ("/certs", "/certificates", "certs"):
        result["open"] = "certificates"
        say("Opening the certificate generator...", "s")
    elif cmd.startswith(("/qr", "qr ")):
        term = raw.split(" ", 1)[1].strip().lower() if " " in raw else ""
        students = store.load_students()
        hit = next((st for st in students
                    if term and term in str(st.get("Name", "")).lower()), None)
        if not hit:
            say("Usage: /qr <student name>", "w")
        else:
            result["qr"] = str(hit.get("UID"))
            say(f"QR code for {hit.get('Name')} \u2014 opening.", "s")
    elif cmd in ("/calendar", "calendar"):
        result["open"] = "calendar"
        say("Opening the attendance calendar...", "s")
    elif cmd in ("/audit", "audit"):
        rows = []
        try:
            if AUDIT_LOG.exists():
                with open(AUDIT_LOG, newline="", encoding="utf-8") as f:
                    rows = list(csv.DictReader(f))
        except Exception:
            pass
        if not rows:
            say("Audit log is empty.", "w")
        else:
            say(f"Last {min(15, len(rows))} of {len(rows)} entries:", "s")
            for r in rows[-15:][::-1]:
                say(f"  {r.get('date','')} {r.get('time','')} | "
                    f"{r.get('name','')} | {r.get('status','')} | "
                    f"+{r.get('xp','0')} XP")
    elif cmd in ("/notifyabsent", "notifyabsent"):
        result["notify_absent"] = True
        say("Sending absence emails to parents...", "w")
    elif cmd in ("/api", "api"):
        on = not bool(settings["api_server"])
        settings.save({"api_server": on})
        if on:
            ok = api_bridge.start()
            say(f"External API {'listening on port ' + str(API_PORT) if ok else 'failed to start'}.",
                "s" if ok else "e")
        else:
            api_bridge.stop()
            say("External API stopped.", "w")
    elif cmd in ("/theme", "theme"):
        result["toggle_theme"] = True
        say("Theme toggled.", "s")
    elif cmd in ("/restart", "restart"):
        result["reload"] = True
        say("Reloading the interface...", "w")
    elif cmd in ("/backup", "backup"):
        BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
        out = BACKUPS_DIR / f"backup_{datetime.now():%Y%m%d_%H%M%S}"
        out.mkdir(exist_ok=True)
        try:
            if EXCEL_FILE.exists():
                shutil.copy2(EXCEL_FILE, out / EXCEL_FILE.name)
            if FACES_DIR.exists():
                shutil.copytree(FACES_DIR, out / "faces", dirs_exist_ok=True)
            if SETTINGS_FILE.exists():
                shutil.copy2(SETTINGS_FILE, out / "settings.json")
            say(f"Backup saved: {out.name}", "s")
        except Exception as e:
            say(f"Backup failed: {e}", "e")
    elif cmd in ("/export", "export"):
        EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
        out = EXPORTS_DIR / f"FaceAttend_Export_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        try:
            shutil.copy2(EXCEL_FILE, out)
            say(f"Roster exported: {out.name}", "s")
        except Exception as e:
            say(f"Export failed: {e}", "e")
    elif raw.strip().lower() in ("/devmode", "/unlock"):
        # Intentionally missing from /help — this is a developer door.
        result["unlock_dev"] = True
        say("Developer tools unlocked. Ctrl+Shift+G for the arcade, "
            "Ctrl+Shift+L to lock again.", "s")
    elif cmd in ("/clear", "clear"):
        result["clear"] = True
        say("console buffer cleared", "w")
    elif cmd in ("/add", "add"):
        result["navigate"] = "home"
        say("Opening dashboard — use Add to enrol a student.", "i")
    elif cmd.startswith("/launch") or cmd.startswith("launch"):
        game = re.sub(r"^/?launch\s*", "", cmd).strip().replace(" ", "").replace("-", "")
        games = {"snake", "trex", "pac", "chess", "liar"}
        if not game:
            say("Usage: /launch <snake|trex|pac|chess|liar>", "w")
        elif game not in games:
            say(f"Unknown game '{game}'. Try /? for the list.", "e")
        else:
            result["navigate"] = f"game:{game}"
            say(f"Launching {game}…", "s")
    elif "analytics" in cmd or cmd in ("/houses", "houses"):
        result["navigate"] = "analytics"
        say("Analytics workspace opened.", "i")
    elif "notepad" in cmd or "whiteboard" in cmd:
        result["navigate"] = "notes"
        say("Notepad workspace opened.", "i")
    else:
        say(f"Unknown command: {raw}", "e")
        say("Type /? to see available commands.", "w")
    return jsonify(result)


# ── recognition self-test ─────────────────────────────────────────
@app.route("/api/recognition/test")
def api_recognition_test():
    """
    Score the enrolled faces against each other.

    For every stored encoding this measures the distance to the same
    student's other photos (should be small) and to every other student
    (should be large). It reports the margin between them and warns about
    any pair that sits close enough to be confused at the current threshold.

    This is the check that tells you whether the system will actually work
    on the day, rather than finding out in front of an audience.
    """
    if np is None:
        return jsonify({"error": "numpy is not installed."}), 500
    uids, vecs = scanner._flat_encodings()
    if len(vecs) < 2:
        return jsonify({"ok": False,
                        "error": "Enrol at least two students with photos first."}), 200

    names = {str(x.get("UID")): x.get("Name") for x in store.load_students()}
    thr = safe_float(settings["threshold"], 0.45)

    same, diff = [], []
    risky_map: dict = {}
    per_student: dict[str, list] = {}
    for u_, v in zip(uids, vecs):
        per_student.setdefault(u_, []).append(v)

    for i in range(len(vecs)):
        for j in range(i + 1, len(vecs)):
            dist = float(np.linalg.norm(np.array(vecs[i]) - np.array(vecs[j])))
            if uids[i] == uids[j]:
                same.append(dist)
            else:
                diff.append(dist)
                if dist <= thr:
                    # Keep one entry per pair of students, not one per pair of
                    # photos — otherwise six photos each produces nine
                    # identical-looking warnings for the same two people.
                    key = tuple(sorted((uids[i], uids[j])))
                    prev = risky_map.get(key)
                    if prev is None or dist < prev["distance"]:
                        risky_map[key] = {
                            "a": names.get(key[0], key[0]),
                            "b": names.get(key[1], key[1]),
                            "distance": round(dist, 3),
                        }

    avg = lambda xs: round(sum(xs) / len(xs), 3) if xs else None
    worst_same = round(max(same), 3) if same else None
    best_diff = round(min(diff), 3) if diff else None
    margin = round(best_diff - worst_same, 3) if (same and diff) else None

    thin = [{"uid": u_, "name": names.get(u_, u_), "photos": len(v)}
            for u_, v in per_student.items() if len(v) < 3]

    risky = sorted(risky_map.values(), key=lambda r: r["distance"])
    if risky:
        verdict = "At risk — some students look similar enough to be confused."
    elif margin is not None and margin < 0.08:
        verdict = "Tight — recognition should work, but the margin is narrow."
    elif thin:
        verdict = "Workable, but several students have too few photos."
    else:
        verdict = "Good separation between students."

    return jsonify({
        "ok": True,
        "students": len(per_student), "encodings": len(vecs),
        "threshold": thr,
        "same_avg": avg(same), "same_worst": worst_same,
        "diff_avg": avg(diff), "diff_closest": best_diff,
        "margin": margin, "risky": risky[:10], "thin": thin,
        "verdict": verdict,
    })


# ── late-arrival slips ────────────────────────────────────────────
@app.route("/api/slip", methods=["POST"])
def api_slip():
    d = request.json or {}
    uid = str(d.get("uid", ""))
    if not may_touch(uid):
        return jsonify({"error": "That student is not in your classes."}), 403
    st = next((x for x in store.load_students() if str(x.get("UID")) == uid), None)
    if not st:
        return jsonify({"error": "Student not found."}), 404
    SLIPS_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w\s-]", "", str(st.get("Name") or uid)).strip() or uid
    path = SLIPS_DIR / f"{safe}_{datetime.now():%Y%m%d_%H%M}.pdf"
    ok = generate_late_slip(path, st.get("Name"), st.get("Roll No"),
                            st.get("Grade"),
                            datetime.now().strftime("%d %B %Y, %H:%M"),
                            d.get("status") or attendance_status()["label"],
                            d.get("reason", ""), settings["school_name"])
    if not ok:
        return jsonify({"error": "reportlab is not installed — "
                                 "run: pip install reportlab"}), 500
    return jsonify({"ok": True, "file": path.name, "folder": str(SLIPS_DIR)})


@app.route("/api/slip/<path:name>")
def api_slip_file(name):
    f = SLIPS_DIR / name
    if f.exists() and f.parent == SLIPS_DIR:
        return send_file(f)
    return ("", 404)


# ── low attendance alerts ─────────────────────────────────────────
@app.route("/api/alerts")
def api_alerts():
    """
    Students below the attendance threshold.

    Chronic absence is what schools actually chase; notifying only about
    today misses the student who is away one day in four.
    """
    limit = safe_float(request.args.get("pct") or settings["low_attendance_pct"], 85.0)
    data = api_report().get_json()
    at_risk = [r for r in data["students"] if r["pct"] < limit]
    at_risk.sort(key=lambda r: r["pct"])
    return jsonify({"threshold": limit, "school_days": data["school_days"],
                    "students": at_risk, "total": len(data["students"])})


@app.route("/api/alerts/notify", methods=["POST"])
def api_alerts_notify():
    """Email the parents of everyone below the threshold."""
    import smtplib
    from email.message import EmailMessage
    user, pw = settings["smtp_user"], settings["smtp_pass"]
    if not user or not pw:
        return jsonify({"error": "Set your SMTP details in Settings first."}), 400
    limit = safe_float((request.json or {}).get("pct")
                       or settings["low_attendance_pct"], 85.0)
    data = api_report().get_json()
    at_risk = [r for r in data["students"] if r["pct"] < limit]
    roster = {str(x.get("UID")): x for x in store.load_students()}
    school = settings["school_name"] or "School"
    sent, skipped = [], []
    try:
        with smtplib.SMTP(settings["smtp_host"],
                          safe_int(settings["smtp_port"], 587), timeout=25) as srv:
            srv.starttls()
            srv.login(user, pw)
            for r in at_risk:
                st = roster.get(r["uid"], {})
                to = str(st.get("Parent Email") or "").strip()
                if not to:
                    skipped.append(r["name"])
                    continue
                m = EmailMessage()
                m["From"], m["To"] = user, to
                m["Subject"] = f"Attendance concern — {r['name']}"
                m.set_content(
                    f"Dear Parent,\n\n{r['name']}'s attendance is currently "
                    f"{r['pct']}%, which is below the {limit:.0f}% we ask for.\n"
                    f"They have attended {r['days_present']} of "
                    f"{data['school_days']} school days.\n\n"
                    f"Please get in touch if there is something we should know "
                    f"about.\n\nKind regards,\n{school}")
                srv.send_message(m)
                sent.append(r["name"])
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return jsonify({"ok": True, "sent": sent, "skipped": skipped})


# ── bulk photo enrolment ──────────────────────────────────────────
@app.route("/api/photos/bulk", methods=["POST"])
def api_photos_bulk():
    """
    Enrol faces from a folder of images named by roll number or UID.

    Photographing 300 students one at a time through the webcam is not
    realistic; a school will already have a photo folder.
    """
    denied = require_admin()
    if denied:
        return denied
    if face_recognition is None:
        return jsonify({"error": "face_recognition is not installed."}), 500

    files = request.files.getlist("photos")
    if not files:
        return jsonify({"error": "No images uploaded."}), 400

    students = store.load_students()
    by_roll = {str(s.get("Roll No") or "").strip().lower(): s for s in students}
    by_uid = {str(s.get("UID") or "").strip().lower(): s for s in students}
    by_name = {str(s.get("Name") or "").strip().lower(): s for s in students}

    matched, no_match, no_face = [], [], []
    FACES_DIR.mkdir(parents=True, exist_ok=True)

    for f in files[:400]:
        stem = Path(f.filename or "").stem.strip()
        key = stem.lower()
        st = by_roll.get(key) or by_uid.get(key) or by_name.get(key)
        if st is None:
            # allow "8C01_aarav.jpg" style names
            head = key.split("_")[0].split("-")[0]
            st = by_roll.get(head) or by_uid.get(head)
        if st is None:
            no_match.append(f.filename)
            continue
        uid = str(st.get("UID"))
        try:
            existing = len(scanner.encodings.get(uid) or [])
            dest = FACES_DIR / (f"{uid}.jpg" if existing == 0
                                else f"{uid}_{existing}.jpg")
            f.save(dest)
            if scanner.learn(uid, dest, append=(existing > 0)):
                matched.append({"name": st.get("Name"), "file": f.filename})
            else:
                no_face.append(f.filename)
        except Exception:
            no_face.append(f.filename)

    return jsonify({"ok": True, "added": len(matched), "matched": matched[:50],
                    "no_match": no_match[:30], "no_face": no_face[:30]})


# ── parent portal ─────────────────────────────────────────────────
@app.route("/portal")
def portal_page():
    return Response(PORTAL_HTML, mimetype="text/html")


@app.route("/api/portal/<uid>")
def api_portal(uid):
    """
    Read-only view of one student, for a parent.

    Deliberately narrow: it returns that child and nothing else, so a
    guessed UID cannot be used to browse the school.
    """
    uid = str(uid).strip().upper()
    st = next((x for x in store.load_students()
               if str(x.get("UID")).upper() == uid), None)
    if not st:
        return jsonify({"error": "No student with that code."}), 404

    rows = [r for r in audit_rows() if str(r.get("uid")).upper() == uid]
    present_days = sorted({r["date"] for r in rows
                           if r.get("status") in ("ontime", "present", "late", "half_day")})
    absences = [{"date": r.get("date"), "reason": r.get("reason") or "not given"}
                for r in rows if r.get("status") in ("absence", "absent")][-12:]
    data = api_report().get_json()
    mine = next((r for r in data["students"] if r["uid"] == uid), None)

    return jsonify({
        "name": st.get("Name"), "grade": st.get("Grade"), "house": st.get("House"),
        "days_present": len(present_days),
        "school_days": data["school_days"],
        "pct": mine["pct"] if mine else 0.0,
        "days_late": mine["days_late"] if mine else 0,
        "xp": safe_int(st.get("XP")), "streak": safe_int(st.get("Streak")),
        "badges": [{"name": b["name"], "icon": b["icon"]}
                   for b in badges.for_student(uid)],
        "absences": absences,
        "present_today": uid in {u.upper() for u in scanner.scanned},
        "school": settings["school_name"] or "School",
    })


# ── exhibition mode & backups ─────────────────────────────────────
@app.route("/api/demo")
def api_demo_status():
    return jsonify({"active": demo.active(), "info": demo.info(),
                    "backups": backups.status()})


@app.route("/api/demo/start", methods=["POST"])
def api_demo_start():
    denied = require_admin()
    if denied:
        return denied
    me = current_user()
    res = demo.start((me or {}).get("username", ""))
    return jsonify(res), (200 if res.get("ok") else 500)


@app.route("/api/demo/restore", methods=["POST"])
def api_demo_restore():
    denied = require_admin()
    if denied:
        return denied
    res = demo.restore()
    return jsonify(res), (200 if res.get("ok") else 400)


@app.route("/api/demo/clear-today", methods=["POST"])
def api_demo_clear():
    return jsonify(demo.clear_today())


@app.route("/api/backup/auto", methods=["POST"])
def api_backup_auto():
    denied = require_admin()
    if denied:
        return denied
    res = backups.run()
    return jsonify(res), (200 if res.get("ok") else 500)


# ── staff accounts (admin only) ───────────────────────────────────
@app.route("/api/users")
def api_users():
    denied = require_admin()
    if denied:
        return denied
    grades = sorted({str(s.get("Grade") or "").strip()
                     for s in store.load_students() if s.get("Grade")})
    return jsonify({"users": accounts.list_users(), "grades": grades})


@app.route("/api/users", methods=["POST"])
def api_users_create():
    denied = require_admin()
    if denied:
        return denied
    d = request.json or {}
    ok, err = accounts.register(d.get("username", ""), d.get("password", ""),
                                d.get("display", ""))
    if not ok:
        return jsonify({"error": err}), 400
    uname = d.get("username", "").strip().lower()
    if d.get("role") in ("admin", "teacher"):
        accounts.set_role(uname, d["role"])
    if isinstance(d.get("classes"), list):
        accounts.set_classes(uname, d["classes"])
    return jsonify({"ok": True})


@app.route("/api/users/<username>", methods=["PUT"])
def api_users_update(username):
    denied = require_admin()
    if denied:
        return denied
    d = request.json or {}
    if "role" in d:
        ok, err = accounts.set_role(username, d["role"])
        if not ok:
            return jsonify({"error": err}), 400
    if isinstance(d.get("classes"), list):
        accounts.set_classes(username, d["classes"])
    if d.get("password"):
        ok, err = accounts.reset_password(username, d["password"])
        if not ok:
            return jsonify({"error": err}), 400
    return jsonify({"ok": True})


@app.route("/api/users/<username>", methods=["DELETE"])
def api_users_delete(username):
    denied = require_admin()
    if denied:
        return denied
    me = current_user()
    if me and me.get("username") == username.strip().lower():
        return jsonify({"error": "You cannot delete your own account."}), 400
    ok, err = accounts.delete_user(username)
    if not ok:
        return jsonify({"error": err}), 400
    return jsonify({"ok": True})


# ── timetable ─────────────────────────────────────────────────────
@app.route("/api/schedule")
def api_schedule():
    return jsonify({"schedule": settings["schedule"],
                    "now": attendance_status(),
                    "today": day_schedule(),
                    "register": register.status()})


@app.route("/api/schedule", methods=["POST"])
def api_schedule_save():
    denied = require_admin()
    if denied:
        return denied
    d = request.json or {}
    sched = dict(settings["schedule"] or {})
    if "enforce" in d:
        sched["enforce"] = bool(d["enforce"])
    if isinstance(d.get("days"), dict):
        days = dict(sched.get("days") or {})
        for k, v in d["days"].items():
            if k not in WEEKDAY_KEYS or not isinstance(v, dict):
                continue
            cur = dict(days.get(k) or {})
            cur["active"] = bool(v.get("active", cur.get("active", True)))
            for f, dflt in (("start", "07:30"), ("late", "08:30"),
                            ("half", "11:00"), ("close", "15:30")):
                hh, mm = _hhmm(v.get(f, cur.get(f, dflt)), dflt)
                cur[f] = f"{hh:02d}:{mm:02d}"
            days[k] = cur
        sched["days"] = days
    if isinstance(d.get("holidays"), list):
        sched["holidays"] = sorted({str(x)[:10] for x in d["holidays"] if x})[:400]
    settings.save({"schedule": sched})
    return jsonify({"ok": True, "schedule": sched, "now": attendance_status()})


@app.route("/api/register/close", methods=["POST"])
def api_register_close():
    user = accounts.session_user(_token())
    res = register.close_now((user or {}).get("username", "manual"))
    return jsonify(res)


# ── cameras ───────────────────────────────────────────────────────
@app.route("/api/cameras")
def api_cameras():
    return jsonify({"cameras": settings["cameras"] or [],
                    "active": scanner.active_source,
                    "running": scanner.running})


@app.route("/api/cameras", methods=["POST"])
def api_cameras_save():
    denied = require_admin()
    if denied:
        return denied
    cams = (request.json or {}).get("cameras", [])
    clean = []
    for c in cams[:8]:
        src = str(c.get("source", "")).strip()
        if not src:
            continue
        clean.append({"name": str(c.get("name", "")).strip() or "Camera",
                      "source": src})
    settings.save({"cameras": clean})
    return jsonify({"ok": True, "cameras": clean})


@app.route("/api/cameras/test", methods=["POST"])
def api_cameras_test():
    """Try a source and report back, so a wrong RTSP URL is caught here."""
    src = str((request.json or {}).get("source", "")).strip()
    if not src:
        return jsonify({"error": "No source given."}), 400
    if cv2 is None:
        return jsonify({"error": "opencv-python is not installed."}), 500
    cap = scanner.open_stream(src)
    if cap is None:
        return jsonify({"ok": False,
                        "error": "Could not open that source. Check the URL, "
                                 "username and password, and that the camera "
                                 "is on the same network."}), 200
    ok, frame = cap.read()
    w = h = 0
    if ok and frame is not None:
        h, w = frame.shape[:2]
    try:
        cap.release()
    except Exception:
        pass
    return jsonify({"ok": bool(ok), "width": w, "height": h,
                    "error": "" if ok else "Opened, but no frame arrived."})


# ── correcting attendance ─────────────────────────────────────────
ABSENCE_REASONS = ["Sick", "Excused", "Late with note", "Medical",
                   "Family", "Marked in error", "Other"]


@app.route("/api/unmark", methods=["POST"])
def api_unmark():
    """
    Undo an attendance mark. Every real register needs this — a mis-scan or
    a wrong click currently could not be corrected at all.
    The XP, day count and streak awarded for that mark are all reversed.
    """
    d = request.json or {}
    uid = str(d.get("uid", ""))
    if not may_touch(uid):
        return jsonify({"error": "That student is not in your classes."}), 403
    reason = str(d.get("reason", "Marked in error")).strip()
    if uid not in scanner.scanned:
        return jsonify({"error": "That student is not marked present today."}), 400

    students = store.load_students()
    target = None
    for st in students:
        if str(st.get("UID")) == uid:
            xp_back = safe_int(settings["xp_ontime"], 5)
            st["No. of Days Present"] = max(0, safe_int(st.get("No. of Days Present")) - 1)
            st["XP"] = max(0, safe_int(st.get("XP")) - xp_back)
            st["Streak"] = max(0, safe_int(st.get("Streak")) - 1)
            target = st
    if target is None:
        return jsonify({"error": "Student not found."}), 404

    store.write_all(students)
    scanner.scanned.discard(uid)
    scanner.save_scans()
    user = accounts.session_user(_token())
    scanner._audit(uid, target.get("Name", uid), False, 0,
                   action="unmark", reason=reason,
                   actor=(user or {}).get("username", ""))
    return jsonify({"ok": True, "name": target.get("Name")})


@app.route("/api/absence", methods=["POST"])
def api_absence():
    """Record WHY a student is absent, without marking them present."""
    d = request.json or {}
    uid = str(d.get("uid", ""))
    reason = str(d.get("reason", "")).strip()
    if not reason:
        return jsonify({"error": "Pick a reason."}), 400
    students = {str(x.get("UID")): x for x in store.load_students()}
    st = students.get(uid)
    if not st:
        return jsonify({"error": "Student not found."}), 404
    user = accounts.session_user(_token())
    scanner._audit(uid, st.get("Name", uid), False, 0,
                   action="absence", reason=reason,
                   actor=(user or {}).get("username", ""))
    return jsonify({"ok": True, "name": st.get("Name"), "reason": reason})


@app.route("/api/audit/verify")
def api_audit_verify():
    return jsonify(audit_verify())


# ── reports ───────────────────────────────────────────────────────
@app.route("/api/report")
def api_report():
    """
    Attendance summary a school office can actually use: per-student
    percentage over a date range, plus absence reasons recorded.
    """
    frm = request.args.get("from", "")
    to = request.args.get("to", "")
    rows = [r for r in audit_rows()
            if r.get("status") in ("ontime", "present", "late", "half_day")]
    if frm:
        rows = [r for r in rows if (r.get("date") or "") >= frm]
    if to:
        rows = [r for r in rows if (r.get("date") or "") <= to]

    observed = sorted({r.get("date") for r in rows if r.get("date")})
    # Prefer the timetable, so weekends and holidays never dilute a
    # percentage. Fall back to observed days when no schedule is enforced.
    scheduled = working_days(frm or (observed[0] if observed else ""), to)
    days = scheduled or observed
    per: dict[str, dict] = {}
    for r in rows:
        u = r.get("uid")
        rec = per.setdefault(u, {"present": set(), "late": 0, "half": 0})
        rec["present"].add(r.get("date"))
        if r.get("status") == "late":
            rec["late"] += 1
        if r.get("status") == "half_day":
            rec["half"] += 1

    reasons: dict[str, list] = {}
    for r in audit_rows():
        if r.get("status") == "absence":
            reasons.setdefault(r.get("uid"), []).append(
                {"date": r.get("date"), "reason": r.get("reason")})

    out = []
    total_days = len(days)
    for st in scope_students(store.load_students()):
        uid = str(st.get("UID"))
        rec = per.get(uid, {"present": set(), "late": 0, "half": 0})
        n = len(rec["present"])
        out.append({
            "uid": uid, "name": st.get("Name"), "grade": st.get("Grade"),
            "house": st.get("House"), "roll": st.get("Roll No"),
            "days_present": n, "days_late": rec["late"],
            "days_half": rec.get("half", 0),
            "pct": round(n / total_days * 100, 1) if total_days else 0.0,
            "absences": reasons.get(uid, []),
        })
    out.sort(key=lambda x: -x["pct"])
    return jsonify({"students": out, "school_days": total_days,
                    "counting": "timetable" if scheduled else "days with scans",
                    "from": days[0] if days else "", "to": days[-1] if days else "",
                    "average": round(sum(x["pct"] for x in out) / len(out), 1) if out else 0.0})


@app.route("/api/report/export", methods=["POST"])
def api_report_export():
    """Write the summary to its own spreadsheet."""
    if Workbook is None:
        return jsonify({"error": "openpyxl is not installed."}), 500
    data = api_report().get_json()
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    out = EXPORTS_DIR / f"Attendance_Report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    hdr = ["Roll No", "Name", "Grade", "House", "Days present",
           "Days late", "Attendance %", "Absence reasons"]
    ws.append(hdr)
    for i in range(1, len(hdr) + 1):
        c = ws.cell(1, i)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="2F6BFF")
        ws.column_dimensions[get_column_letter(i)].width = max(13, len(hdr[i-1]) + 4)
    for r in data["students"]:
        ws.append([r["roll"], r["name"], r["grade"], r["house"],
                   r["days_present"], r["days_late"], r["pct"],
                   "; ".join(f"{a['date']}: {a['reason']}" for a in r["absences"])])
    ws.freeze_panes = "A2"
    try:
        wb.save(out)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return jsonify({"ok": True, "path": str(out), "file": out.name,
                    "rows": len(data["students"])})


# ── badges ────────────────────────────────────────────────────────
@app.route("/api/badges")
def api_badges():
    students = scope_students(store.load_students())
    out = []
    for st in students:
        uid = str(st.get("UID"))
        out.append({
            "uid": uid,
            "name": st.get("Name"),
            "house": st.get("House"),
            "grade": st.get("Grade"),
            "badges": badges.all_with_progress(st),
            "earned": len(badges.for_student(uid)),
        })
    return jsonify({"students": out, "defs": Badges.DEFS})


# ── certificates ──────────────────────────────────────────────────
@app.route("/api/certificates", methods=["POST"])
def api_certificates():
    d = request.json or {}
    min_pct = safe_float(d.get("min_pct", settings["cert_min_pct"]), 90.0)
    only = set(str(u) for u in (d.get("uids") or []))
    students = store.load_students()
    total_days = max([safe_int(s.get("No. of Days Present")) for s in students] or [0])
    if total_days <= 0:
        return jsonify({"error": "No attendance recorded yet."}), 400

    CERTS_DIR.mkdir(parents=True, exist_ok=True)
    made, skipped = [], []
    for st in students:
        uid = str(st.get("UID"))
        if only and uid not in only:
            continue
        days = safe_int(st.get("No. of Days Present"))
        pct = days / total_days * 100 if total_days else 0
        if pct < min_pct:
            skipped.append({"name": st.get("Name"), "pct": round(pct, 1)})
            continue
        safe_name = re.sub(r"[^\w\s-]", "", str(st.get("Name") or uid)).strip()
        path = CERTS_DIR / f"{safe_name or uid}_{uid}.pdf"
        ok = generate_certificate(path, st.get("Name"), st.get("Roll No") or "-",
                                  st.get("Grade") or "", pct, days,
                                  settings["school_name"])
        if ok:
            made.append({"name": st.get("Name"), "pct": round(pct, 1),
                         "file": path.name})
        else:
            return jsonify({"error": "reportlab is not installed — "
                                     "run: pip install reportlab"}), 500
    return jsonify({"ok": True, "made": made, "skipped": skipped,
                    "folder": str(CERTS_DIR)})


@app.route("/api/certificates/<path:name>")
def api_cert_file(name):
    f = CERTS_DIR / name
    if f.exists() and f.parent == CERTS_DIR:
        return send_file(f)
    return ("", 404)


# ── QR codes ──────────────────────────────────────────────────────
@app.route("/api/qr/<uid>")
def api_qr(uid):
    students = {str(s["UID"]): s for s in store.load_students()}
    st = students.get(str(uid))
    if not st:
        return jsonify({"error": "Student not found"}), 404
    path = generate_student_qr(uid, str(st.get("Name") or ""),
                               str(st.get("Roll No") or ""))
    if path is None:
        return jsonify({"error": "qrcode not installed — pip install qrcode"}), 500
    return send_file(path, mimetype="image/png")


# ── calendar heatmap ──────────────────────────────────────────────
@app.route("/api/calendar")
def api_calendar():
    now = datetime.now()
    y = safe_int(request.args.get("year"), now.year)
    m = safe_int(request.args.get("month"), now.month)
    m = min(12, max(1, m))
    return jsonify(calendar_data(y, m, len(store.load_students())))


# ── audit log ─────────────────────────────────────────────────────
@app.route("/api/audit")
def api_audit():
    rows = []
    try:
        if AUDIT_LOG.exists():
            with open(AUDIT_LOG, newline="", encoding="utf-8") as f:
                rows = list(csv.DictReader(f))
    except Exception:
        pass
    return jsonify({"rows": rows[-400:][::-1], "total": len(rows)})


# ── bulk import via pasted text ───────────────────────────────────
@app.route("/api/bulk", methods=["POST"])
def api_bulk():
    """
    Accepts pasted lines: Name, Roll, Grade, House, Phone, Email
    One student per line. Blank fields are fine.
    """
    text = (request.json or {}).get("text", "")
    added, failed = 0, []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.lower().startswith("name"):
            continue
        parts = [x.strip() for x in re.split(r"[,\t;]", line)]
        if not parts or not parts[0]:
            continue
        try:
            store.add({
                "Name": parts[0],
                "Roll No": parts[1] if len(parts) > 1 else "",
                "Grade": parts[2] if len(parts) > 2 else "",
                "House": (parts[3] if len(parts) > 3 and parts[3] in HOUSES else "Blue"),
                "Parent Phone": parts[4] if len(parts) > 4 else "",
                "Parent Email": parts[5] if len(parts) > 5 else "",
            })
            added += 1
        except Exception as e:
            failed.append(f"{parts[0]}: {e}")
    return jsonify({"ok": True, "added": added, "failed": failed})


# ── absent notifications ──────────────────────────────────────────
@app.route("/api/notify-status")
def api_notify_status():
    return jsonify(notifier.status())


@app.route("/api/notify-absent", methods=["POST"])
def api_notify_absent():
    res = notifier.send_now()
    if not res.get("ok"):
        return jsonify({"error": res.get("error", "Send failed.")}), 400
    notifier.last_sent_date = str(date.today())
    notifier.last_result = f"{len(res['sent'])} sent (manual)"
    notifier._save()
    return jsonify({"ok": True, "sent": res["sent"], "skipped": res["skipped"]})


# ── encryption ────────────────────────────────────────────────────
@app.route("/api/vault", methods=["POST"])
def api_vault():
    action = (request.json or {}).get("action", "status")
    if not vault.available():
        return jsonify({"error": "cryptography not installed — "
                                 "pip install cryptography"}), 500
    if action == "encrypt":
        ok = vault.encrypt_file(EXCEL_FILE)
        return jsonify({"ok": ok, "file": str(EXCEL_FILE) + ".enc"})
    if action == "decrypt":
        enc = Path(str(EXCEL_FILE) + ".enc")
        ok = vault.decrypt_file(enc)
        return jsonify({"ok": ok})
    return jsonify({"available": True,
                    "encrypted": Path(str(EXCEL_FILE) + ".enc").exists()})


# ── section / grade change in bulk ────────────────────────────────
@app.route("/api/section", methods=["POST"])
def api_section():
    denied = require_admin()
    if denied:
        return denied
    d = request.json or {}
    uids = set(str(u) for u in (d.get("uids") or []))
    grade = str(d.get("grade", "")).strip()
    house = str(d.get("house", "")).strip()
    if not uids or (not grade and not house):
        return jsonify({"error": "Pick students and a new grade or house."}), 400
    students = store.load_students()
    n = 0
    for st in students:
        if str(st.get("UID")) in uids:
            if grade:
                st["Grade"] = grade
            if house in HOUSES:
                st["House"] = house
            n += 1
    store.write_all(students)
    return jsonify({"ok": True, "changed": n})


# ── import / export / backup ──────────────────────────────────────
@app.route("/api/export", methods=["POST"])
def api_export():
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    out = EXPORTS_DIR / f"FaceAttend_Export_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    try:
        shutil.copy2(EXCEL_FILE, out)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return jsonify({"ok": True, "path": str(out)})


@app.route("/api/backup", methods=["POST"])
def api_backup():
    BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = f"{datetime.now():%Y%m%d_%H%M%S}"
    out = BACKUPS_DIR / f"backup_{stamp}"
    out.mkdir(exist_ok=True)
    try:
        if EXCEL_FILE.exists():
            shutil.copy2(EXCEL_FILE, out / EXCEL_FILE.name)
        if FACES_DIR.exists():
            shutil.copytree(FACES_DIR, out / "faces", dirs_exist_ok=True)
        if SETTINGS_FILE.exists():
            shutil.copy2(SETTINGS_FILE, out / "settings.json")
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return jsonify({"ok": True, "path": str(out)})


@app.route("/api/import", methods=["POST"])
def api_import():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file"}), 400
    added = 0
    try:
        if f.filename.lower().endswith(".csv"):
            rows = list(csv.DictReader(io.StringIO(f.read().decode("utf-8", "ignore"))))
        else:
            wb = load_workbook(io.BytesIO(f.read()), data_only=True)
            ws = wb.active
            data = list(ws.iter_rows(values_only=True))
            hdr = [str(h or "") for h in data[0]]
            rows = [dict(zip(hdr, r)) for r in data[1:]]
        for r in rows:
            name = str(r.get("Name") or "").strip()
            if not name:
                continue
            store.add({
                "Name": name,
                "Roll No": r.get("Roll No", ""),
                "Grade": r.get("Grade", ""),
                "House": r.get("House") or "Blue",
                "Parent Phone": r.get("Parent Phone", ""),
                "Parent Email": r.get("Parent Email", ""),
            })
            added += 1
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"ok": True, "added": added})


# ── mail ──────────────────────────────────────────────────────────
@app.route("/api/mail", methods=["POST"])
def api_mail():
    import smtplib
    from email.message import EmailMessage
    d = request.json or {}
    user, pw = settings["smtp_user"], settings["smtp_pass"]
    if not user or not pw:
        return jsonify({"error": "Set your SMTP address and app password in Settings first."}), 400
    try:
        m = EmailMessage()
        m["From"], m["To"] = user, d.get("to", "")
        m["Subject"] = d.get("subject", "Message from School")
        m.set_content(d.get("body", ""))
        with smtplib.SMTP(settings["smtp_host"], safe_int(settings["smtp_port"], 587), timeout=20) as srv:
            srv.starttls()
            srv.login(user, pw)
            srv.send_message(m)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return jsonify({"ok": True})


# ── whiteboard + notes ────────────────────────────────────────────
@app.route("/api/board", methods=["POST"])
def api_board():
    img = (request.json or {}).get("image", "")
    if "," not in img:
        return jsonify({"error": "Bad image"}), 400
    BOARD_DIR.mkdir(parents=True, exist_ok=True)
    out = BOARD_DIR / f"board_{datetime.now():%Y%m%d_%H%M%S}.png"
    out.write_bytes(base64.b64decode(img.split(",", 1)[1]))
    return jsonify({"ok": True, "path": str(out)})


@app.route("/api/notes")
def api_notes():
    NOTES_DIR.mkdir(parents=True, exist_ok=True)
    notes = [{"name": p.name, "size": p.stat().st_size}
             for p in sorted(NOTES_DIR.glob("*.txt"), key=lambda p: -p.stat().st_mtime)]
    return jsonify({"notes": notes})


@app.route("/api/notes/<name>")
def api_note_get(name):
    p = NOTES_DIR / name
    if not p.exists() or p.parent != NOTES_DIR:
        return jsonify({"error": "Not found"}), 404
    return jsonify({"body": p.read_text("utf-8", errors="ignore")})


@app.route("/api/notes", methods=["POST"])
def api_note_save():
    d = request.json or {}
    name = re.sub(r"[^\w\s.-]", "", str(d.get("name", "Untitled note")))[:80].strip() or "Untitled note"
    if not name.endswith(".txt"):
        name += ".txt"
    NOTES_DIR.mkdir(parents=True, exist_ok=True)
    (NOTES_DIR / name).write_text(d.get("body", ""), "utf-8")
    return jsonify({"ok": True, "name": name})


@app.route("/api/notes/<name>", methods=["DELETE"])
def api_note_del(name):
    p = NOTES_DIR / name
    if p.exists() and p.parent == NOTES_DIR:
        p.unlink()
        return jsonify({"ok": True})
    return jsonify({"error": "Not found"}), 404


# ══════════════════════════════════════════════════════════════════════════
#  STARTUP
# ══════════════════════════════════════════════════════════════════════════
def free_port(start=8770) -> int:
    for p in range(start, start + 60):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start


def open_browser(url: str):
    """
    Open the app in a browser. Firefox is preferred where we can find it,
    but we always fall back to the system default so this works anywhere.
    """
    time.sleep(1.4)

    if IS_WINDOWS:
        # Try Firefox in its usual install locations, then the default browser.
        candidates = [
            Path(os.getenv("ProgramFiles", r"C:\Program Files")) / "Mozilla Firefox" / "firefox.exe",
            Path(os.getenv("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Mozilla Firefox" / "firefox.exe",
        ]
        for exe in candidates:
            try:
                if exe.exists():
                    subprocess.Popen([str(exe), "-new-window", url])
                    return
            except Exception:
                pass
        try:
            os.startfile(url)          # type: ignore[attr-defined]
            return
        except Exception:
            pass

    elif IS_MAC:
        for cmd in (["open", "-a", "Firefox", url], ["open", url]):
            try:
                subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return
            except FileNotFoundError:
                continue

    else:  # Linux / BSD
        for cmd in (["firefox", "--new-window", url],
                    ["firefox", url],
                    ["xdg-open", url]):
            try:
                subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return
            except FileNotFoundError:
                continue

    try:
        webbrowser.open(url)
    except Exception:
        pass


def _qt_backend_available() -> bool:
    """Is the pip-installable Qt webview backend importable?"""
    try:
        import qtpy  # noqa: F401
        from PySide6.QtWebEngineWidgets import QWebEngineView  # noqa: F401
        return True
    except Exception:
        try:
            from PyQt6.QtWebEngineWidgets import QWebEngineView  # noqa: F401
            return True
        except Exception:
            return False


def _gtk_backend_available() -> bool:
    """
    Is the system GTK webview importable from THIS interpreter?

    This often fails inside a virtualenv: python3-gi is a compiled extension
    built for the system Python, so a venv on a different minor version
    cannot load it. That is exactly why the Qt backend exists as a fallback.
    """
    try:
        import gi
        gi.require_version("Gtk", "3.0")
        for v in ("4.1", "4.0", "6.0"):
            try:
                gi.require_version("WebKit2", v)
                return True
            except ValueError:
                continue
        return False
    except Exception:
        return False


def install_window_backend() -> bool:
    """Install whatever this platform needs for a native window."""
    import subprocess as _sp
    if IS_WINDOWS or IS_MAC:
        pkgs = ["pywebview"]          # the OS provides the webview
        size = "a few MB"
    else:
        pkgs = ["pywebview", "qtpy", "PySide6-Essentials", "PySide6-Addons"]
        size = "about 240 MB"
    print(f"\n  Installing the native-window backend ({size}).")
    print("  One-time download; afterwards FaceAttend opens in its")
    print("  own desktop window instead of a browser.\n")
    cmd = [sys.executable, "-m", "pip", "install", *pkgs]
    in_venv = hasattr(sys, "real_prefix") or sys.base_prefix != sys.prefix
    rc = _sp.call(cmd)
    if rc != 0 and not in_venv:
        rc = _sp.call(cmd + ["--break-system-packages"])
    if rc == 0:
        print("\n  Done. Start FaceAttend again for the app window.\n")
        return True
    print("\n  Install failed. Run this yourself:")
    print(f"      {sys.executable} -m pip install " + " ".join(pkgs) + "\n")
    return False


def run_native_window(url: str, port: int) -> bool:
    """
    Open FaceAttend in a real desktop window via the system webview.
    WebKitGTK or Qt on Linux, WebView2 on Windows, WKWebView on macOS.
    Returns True if a window was shown, False to fall back to a browser.
    """
    import contextlib as _ctx
    import io as _io
    import logging as _lg
    _lg.getLogger("pywebview").setLevel(_lg.CRITICAL)
    _lg.getLogger("webview").setLevel(_lg.CRITICAL)
    _sink = _io.StringIO()

    # Only Linux/BSD needs an explicit toolkit check. Windows and macOS get
    # their webview from the OS, so guarding on GTK/Qt there would wrongly
    # disable the native window.
    if not (IS_WINDOWS or IS_MAC):
        if not (_gtk_backend_available() or _qt_backend_available()):
            return False

        # opencv-python ships bundled Qt5 plugins and points
        # QT_QPA_PLATFORM_PLUGIN_PATH at them as soon as cv2 is imported.
        # That path is incompatible with PySide6 and makes the window fail
        # with "Could not load the Qt platform plugin xcb".
        if _qt_backend_available() and not _gtk_backend_available():
            try:
                import PySide6
                plugins = Path(PySide6.__file__).parent / "Qt" / "plugins"
                if plugins.exists():
                    os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = str(plugins / "platforms")
                    os.environ["QT_PLUGIN_PATH"] = str(plugins)
                else:
                    os.environ.pop("QT_QPA_PLATFORM_PLUGIN_PATH", None)
            except Exception:
                os.environ.pop("QT_QPA_PLATFORM_PLUGIN_PATH", None)

            # QtWebEngine's sandbox refuses to start as root.
            try:
                if hasattr(os, "geteuid") and os.geteuid() == 0:
                    flags = os.environ.get("QTWEBENGINE_CHROMIUM_FLAGS", "")
                    if "--no-sandbox" not in flags:
                        os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = (
                            flags + " --no-sandbox").strip()
            except Exception:
                pass

    try:
        with _ctx.redirect_stderr(_sink), _ctx.redirect_stdout(_sink):
            import webview
    except Exception:
        return False

    try:
        with _ctx.redirect_stderr(_sink), _ctx.redirect_stdout(_sink):
            webview.initialize()
        window = webview.create_window(
            f"FaceAttend {APP_VERSION}",
            url,
            width=1500, height=940,
            min_size=(1024, 680),
            background_color="#05070d",
            text_select=True,
            confirm_close=False,
        )

        def _shutdown(_w=None):
            try:
                scanner.stop()
            except Exception:
                pass
            os._exit(0)

        try:
            window.events.closed += _shutdown
        except Exception:
            pass

        webview.start()
        _shutdown()
        return True
    except Exception as exc:
        first = str(exc).splitlines()[0][:90] if str(exc) else type(exc).__name__
        print(f"  native window failed ({first})")
        if "xcb" in str(exc).lower() or "platform plugin" in str(exc).lower():
            print("  tip: sudo apt install libxcb-cursor0")
        return False


def open_app_window(url: str) -> str:
    """
    Best app-like window without a native webview. Chromium's --app= gives a
    frameless window; Firefox has no equivalent, but --kiosk removes all
    browser chrome, which is the next best thing.
    """
    for exe in ("chromium", "chromium-browser", "google-chrome",
                "google-chrome-stable", "brave-browser", "microsoft-edge"):
        if shutil.which(exe):
            try:
                subprocess.Popen(
                    [exe, f"--app={url}", "--window-size=1500,940",
                     f"--user-data-dir={APP_DIR / '.browser'}"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return "chromium app window"
            except Exception:
                pass
    if shutil.which("firefox"):
        try:
            subprocess.Popen(["firefox", "--kiosk", url],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return "firefox kiosk (press F11 or Ctrl+W to exit)"
        except Exception:
            pass
    open_browser(url)
    return "browser"


def main():
    ensure_dirs()

    # rebuild any face encodings that exist as photos but not in the pickle
    if face_recognition is not None:
        for p in FACES_DIR.glob("*.jpg"):
            uid = p.stem
            if uid not in scanner.encodings:
                scanner.learn(uid, p)

    # --port N lets a launcher pick the port (the .deb launcher does this so
    # it can wait for the server before opening the app window).
    port = None
    if "--port" in sys.argv:
        try:
            port = int(sys.argv[sys.argv.index("--port") + 1])
        except (IndexError, ValueError):
            port = None
    if port is None:
        port = free_port()
    url = f"http://127.0.0.1:{port}"

    # Older Windows consoles are cp1252 and choke on box-drawing characters.
    fancy = True
    try:
        enc = (sys.stdout.encoding or "").lower()
        fancy = ("utf" in enc) or not IS_WINDOWS
        if IS_WINDOWS:
            try:
                sys.stdout.reconfigure(encoding="utf-8")   # type: ignore[attr-defined]
                fancy = True
            except Exception:
                pass
    except Exception:
        fancy = False

    bar, arrow = ("=" * 49, "->") if not fancy else ("\u2550" * 49, "\u2192")
    print()
    print(f"  {bar}")
    print(f"    FaceAttend {APP_VERSION}  -  web edition")
    print(f"  {bar}")
    print(f"  {arrow} open:  {url}")
    print(f"  {arrow} data:  {APP_DIR}")
    if MISSING:
        print(f"  !  optional packages missing: {', '.join(MISSING)}")
        if any(m in MISSING for m in ("face_recognition", "dlib")):
            print("     Face recognition is off; attendance can still be marked")
            print("     by hand. To enable the camera engine (10-25 min build):")
            print("         python FaceAttend.py --setup-face")
    print(f"  {arrow} keep this window open; press Ctrl+C to stop")
    print()

    # Window mode: default is a native app window. --browser forces a normal
    # browser tab; --no-browser opens nothing (used by external launchers).
    want_native = ("--browser" not in sys.argv) and ("--no-browser" not in sys.argv)

    if want_native:
        # Flask runs in the background; the webview owns the main thread,
        # which GUI toolkits require.
        threading.Thread(
            target=lambda: app.run(host="127.0.0.1", port=port,
                                   debug=False, threaded=True,
                                   use_reloader=False),
            daemon=True,
        ).start()

        for _ in range(60):            # wait for the server to answer
            try:
                with socket.create_connection(("127.0.0.1", port), 0.3):
                    break
            except OSError:
                time.sleep(0.15)

        if run_native_window(url, port):
            return

        how = open_app_window(url)
        print(f"  window: {how}")
        if how != "chromium app window":
            print()
            print("  For a proper resizable app window, run once:")
            if IS_WINDOWS or IS_MAC:
                print(f"      {sys.executable} -m pip install pywebview")
            else:
                print(f"      {sys.executable} -m pip install pywebview qtpy "
                      "PySide6-Essentials PySide6-Addons")
            print("  or:  python FaceAttend.py --setup-window")
            print()
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            scanner.stop()
            print("\n  FaceAttend stopped. Bye.")
            return

    if "--no-browser" not in sys.argv:
        threading.Thread(target=open_browser, args=(url,), daemon=True).start()

    try:
        app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
    except KeyboardInterrupt:
        pass
    except Exception as exc:
        print(f"\n  FaceAttend hit an error: {exc}\n")
        if IS_WINDOWS:
            # Double-clicked .py files close instantly; hold the window open
            # so the error is actually readable.
            try:
                input("  Press Enter to close...")
            except Exception:
                pass
        raise
    finally:
        scanner.stop()
        print("\n  FaceAttend stopped. Bye.")


if __name__ == "__main__":
    main()
