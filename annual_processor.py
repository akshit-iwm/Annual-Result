"""
=============================================================================
ANNUAL RESULT PROCESSOR  v2  — API-based download/upload, 24×7 resilient
=============================================================================
Flow:
  1. ACE Annual Poller Thread monitors "Annual Update Ace.xlsx" every 30 min:
     a. Opens the file in a dedicated Excel instance via xlwings
     b. Calls RefreshAll to get fresh ACE data
     c. Reads all Company Names from column B (row 3 onward)
     d. Queues any company not yet processed today (5-min initial delay)
  2. Worker Thread picks items from queue after their delay:
     a. Download Excel from Charcha (API)
     b. Open in xlwings
     c. Find ANN / ANNUAL marker in header row 1
     d. In the ANNUAL section (columns AFTER the marker only):
        - Find MAR-25 column (source)
        - Find / create MAR-26 column (target)
        - COPY MAR-25 formulas → PASTE into MAR-26 column
          (Presentation Data rows are skipped)
     e. RefreshAll + CalculateFull
     f. Compare PAT before vs after
     g. Check Sources of Funds > 0
     h. If both updated → save, upload via API, send Slack
     i. If not yet live  → delete file, re-queue 30 min later

No NSE/BSE scraping.  No database lookups or writes.
No Selenium — all Charcha I/O goes through api_client.APIClient.
=============================================================================
"""

import os
import sys
import re
import time
import json
import logging
import traceback
import psutil
import winreg
import requests
import pandas as pd
from datetime import datetime, timedelta
from queue import PriorityQueue, Empty
from threading import Thread, Lock, Event
import xlwings as xw

from api_client import APIClient


# ================================================================
# PyInstaller compatibility
# ================================================================
def get_base_path():
    """Return the directory of the EXE (frozen) or script (development)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

EXE_DIR = get_base_path()


# ================================================================
# Logging
# ================================================================
LOG_FILE = os.path.join(EXE_DIR, "annual_processor.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)


# ================================================================
# Configuration
# ================================================================

# --- Directories ---
BASE_DIR         = os.environ.get("BASE_DIR", EXE_DIR)
RESULT_FILES_DIR = os.path.join(BASE_DIR, "result_files")
UPDATED_DIR      = os.path.join(BASE_DIR, "Updated_Excel")
DATA_DIR         = os.path.join(BASE_DIR, "data")
SKIPPED_LOG      = os.path.join(BASE_DIR, "skipped_log_annual.txt")

# --- ACE Annual Excel ---
# Must be placed in the same folder as the EXE (or override via env var)
ANNUAL_ACE_EXCEL = os.environ.get(
    "ANNUAL_ACE_EXCEL",
    os.path.join(EXE_DIR, "Annual Update Ace.xlsx"),
)

# --- Charcha companies allow-list ---
# A plain text file (one company name per line, optionally quoted) listing the
# companies that actually exist in the Charcha DB. A company NOT in this list is
# skipped immediately (no download attempt), avoiding the slow search-and-fail.
# If the file is missing, the allow-list check is disabled (all names allowed).
CHARCHA_LIST_FILE = os.environ.get(
    "CHARCHA_LIST_FILE",
    os.path.join(EXE_DIR, "CharchaCompaniesList.txt"),
)

# --- Slack ---
SLACK_PROCESSED_WEBHOOK = (
    "https://hooks.slack.com/services/T7EJU2RAL/B09KT4SSTJN/"
    "8JhPDQztb9RF1XlyZF8Ym18q"
)

# --- Timing ---
# All env-overridable so cadence can be tuned WITHOUT a code change / rebuild.
# For a one-time bulk reprocess of already-published companies you can speed
# things up by lowering ANNUAL_INITIAL_DELAY_MINS (e.g. 0) and EXCEL_REFRESH_WAIT
# (e.g. 15). Leave them at defaults for the normal "watch for fresh results" run.
ANNUAL_POLL_INTERVAL      = int(os.environ.get("ANNUAL_POLL_INTERVAL", "1800"))  # 30 min
ANNUAL_INITIAL_DELAY_MINS = int(os.environ.get("ANNUAL_INITIAL_DELAY_MINS", "1"))
ANNUAL_RETRY_DELAY_MINS   = int(os.environ.get("ANNUAL_RETRY_DELAY_MINS", "1"))
EXCEL_REFRESH_WAIT        = int(os.environ.get("EXCEL_REFRESH_WAIT", "20"))   # sec wait after company Excel refresh
ACE_REFRESH_WAIT          = int(os.environ.get("ACE_REFRESH_WAIT", "300"))    # max sec wait after ACE Excel refresh
COMPANY_SEARCH_TIMEOUT    = int(os.environ.get("COMPANY_SEARCH_TIMEOUT", "8"))  # sec to find a search match before "not found"

# 24×7 resilience: cooldown before restarting the main loop after unhandled error
RUN_LOOP_COOLDOWN         = int(os.environ.get("RUN_LOOP_COOLDOWN", "60"))  # seconds


class CompanyNotFoundError(Exception):
    """Raised when a company name is not in the Charcha allow-list.
    The worker skips fast (marks failed) without attempting a download."""
    pass

for d in [RESULT_FILES_DIR, UPDATED_DIR, DATA_DIR]:
    os.makedirs(d, exist_ok=True)


# ================================================================
# Thread-Safe State
# ================================================================
state_lock = Lock()
done_today = set()
in_queue   = set()
stop_event = Event()

# ================================================================
# API Client (singleton, lazily initialised in run())
# ================================================================
_api_client: APIClient | None = None


# ================================================================
# MONTH ABBREVIATION MAP
# ================================================================
MONTH_MAP = {
    "JAN": 1, "JANUARY": 1,
    "FEB": 2, "FEBRUARY": 2,
    "MAR": 3, "MARCH": 3,
    "APR": 4, "APRIL": 4,
    "MAY": 5,
    "JUN": 6, "JUNE": 6,
    "JUL": 7, "JULY": 7,
    "AUG": 8, "AUGUST": 8,
    "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
    "OCT": 10, "OCTOBER": 10,
    "NOV": 11, "NOVEMBER": 11,
    "DEC": 12, "DECEMBER": 12,
}

MONTH_MAP_INV = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


def parse_header_date(val) -> tuple | None:
    if val is None:
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        m = val.month
        y = val.year % 100
        return (m, y)

    s = str(val).strip().upper().replace(",", "")
    if not s or len(s) < 3:
        return None

    for pattern in [
        r"([A-Z]+)[^A-Z0-9]*(\d{2,4})",
        r"(\d{1,2})[/\-](\d{2,4})",
        r"(\d{4})[/\-](\d{1,2})",
    ]:
        m_obj = re.match(pattern, s)
        if not m_obj:
            continue
        g1, g2 = m_obj.group(1), m_obj.group(2)

        if g1.isalpha():
            month_num = MONTH_MAP.get(g1)
            if month_num is None:
                continue
            year_num = int(g2)
            if year_num > 99:
                year_num = year_num % 100
            return (month_num, year_num)

        if g1.isdigit() and g2.isdigit():
            a, b = int(g1), int(g2)
            if a > 31:
                month_num = b if b <= 12 else None
                year_num  = a % 100
            else:
                month_num = a if a <= 12 else None
                year_num  = b % 100 if b > 0 else None
            if month_num and year_num is not None:
                return (month_num, year_num)

    return None


def date_matches(val, target_month: int, target_yy: int) -> bool:
    parsed = parse_header_date(val)
    if parsed is None:
        return False
    return parsed == (target_month, target_yy)


def is_ann_marker(val) -> bool:
    """Return True if the header cell marks the start of the ANNUAL section."""
    if val is None:
        return False
    v = str(val).strip().upper().replace(" ", "")
    return "ANNUAL" in v or v == "ANN"


# ================================================================
# Utility Helpers
# ================================================================

def clean_company_name(name: str) -> str:
    if not name:
        return ""
    return name.strip().rstrip(".")


# ----------------------------------------------------------------------------
# Charcha companies allow-list
# ----------------------------------------------------------------------------
# Loaded once (lazily) from CHARCHA_LIST_FILE into a normalized set. A company
# whose normalized name is not in this set is skipped without a download attempt.
_charcha_names = None          # set of normalized names, or None if not loaded
_charcha_lock  = Lock()


def _normalize_name(name: str) -> str:
    """Normalize a company name for tolerant matching: strip quotes/space,
    collapse internal whitespace, casefold. e.g. '\"Venky'S (India) Ltd\"' and
    \"venky's (india) ltd\" compare equal."""
    if name is None:
        return ""
    s = str(name).strip().strip('"').strip("'").strip()
    s = s.rstrip(".")
    s = " ".join(s.split())     # collapse runs of whitespace
    return s.casefold()


def _load_charcha_names() -> set:
    """Load and cache the Charcha allow-list. Returns an empty set (meaning
    'allow everything') if the file is absent or unreadable."""
    global _charcha_names
    with _charcha_lock:
        if _charcha_names is not None:
            return _charcha_names
        names = set()
        try:
            if os.path.exists(CHARCHA_LIST_FILE):
                with open(CHARCHA_LIST_FILE, "r", encoding="utf-8") as f:
                    for line in f:
                        norm = _normalize_name(line)
                        if norm:
                            names.add(norm)
                logging.info(
                    f"Charcha allow-list loaded: {len(names)} companies "
                    f"from {CHARCHA_LIST_FILE}"
                )
            else:
                logging.warning(
                    f"Charcha allow-list file not found: {CHARCHA_LIST_FILE}. "
                    f"Allow-list check DISABLED (all companies will be attempted)."
                )
        except Exception as e:
            logging.warning(
                f"Failed to load Charcha allow-list ({e}). Check DISABLED."
            )
        _charcha_names = names
        return _charcha_names


def is_in_charcha_list(company_name: str) -> bool:
    """True if the company is in the Charcha allow-list. If the list is empty
    (file missing / disabled), returns True for everything (fail-open)."""
    names = _load_charcha_names()
    if not names:
        return True   # list disabled -> don't block anything
    return _normalize_name(company_name) in names


def trust_download_folder(folder_path):
    try:
        path    = r"Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
        key     = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
        new_key = winreg.CreateKey(key, "AutomationDownload")
        winreg.SetValueEx(new_key, "Path",            0, winreg.REG_SZ,    folder_path)
        winreg.SetValueEx(new_key, "AllowSubfolders", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(new_key, "Description",     0, winreg.REG_SZ,    "Automated Downloads")
        logging.info(f"Trusted folder registered: {folder_path}")
    except Exception as e:
        logging.warning(f"Could not trust folder: {e}")


def force_kill_excel():
    for proc in psutil.process_iter(["name"]):
        try:
            if proc.info["name"] and "EXCEL.EXE" in proc.info["name"].upper():
                proc.kill()
        except Exception:
            pass
    time.sleep(1)


# ----------------------------------------------------------------------------
# ACE Add-in loader
# ----------------------------------------------------------------------------
# The ACE/Accord data plug-in is TWO components that BOTH must be loaded:
#   1. ACEEQ_XL_NXT64.xll  (Excel Add-in)  -> provides the worksheet UDFs that
#      populate the 'Annual Raw' sheet. Missing -> every data cell = #NAME?.
#   2. ACEEQ_XL_NXT        (COM Add-in)    -> provides the ribbon + data-refresh
#      plumbing. Missing -> RefreshAll pulls no new rows (stale company list).
#
# Excel instances created via automation (xw.App) do NOT auto-load either one,
# even though a manual double-click does. We must load both explicitly before
# any RefreshAll / recalculation, or the workbook gets corrupted on save.
#
# Override the .xll path via env var ACE_XLL_PATH if the install dir differs.
ACE_XLL_PATH = os.environ.get(
    "ACE_XLL_PATH",
    r"E:\Ace Equity Nxt\ActiveXl\ACEEQ_XL_NXT64.xll",
)
ACE_COM_PROGID = os.environ.get("ACE_COM_PROGID", "ACEEQ_XL_NXT")

# --- ACE ribbon refresh via keyboard accelerators (KeyTips) -----------------
# The ACE 'Refresh' command is a custom Excel-DNA ribbon button whose onAction
# callback is compiled into ACEEQ_XL_NXT.dll (no .dna on disk, not macro-
# callable, no scriptable COM .Object). The only automation-friendly way to
# invoke it is to replay the ribbon keyboard accelerators a user would press:
#   Alt -> tab KeyTip "Y2" (ACEEQ XL NXT tab) -> button KeyTip "Y8" (Refresh).
# Both KeyTips are read off the Excel UI (press Alt) and overridable via env
# vars so adjusting them never needs a code change / rebuild.
ACE_RIBBON_TAB_KEYS    = os.environ.get("ACE_RIBBON_TAB_KEYS", "Y2")
ACE_RIBBON_REFRESH_KEY = os.environ.get("ACE_RIBBON_REFRESH_KEY", "Y8")


def _force_excel_foreground(app) -> bool:
    """
    Bring the Excel main window to the OS foreground using Win32 so SendKeys
    lands in Excel, not the console. Excel's own ActiveWindow.Activate() does
    NOT do this (it only activates a workbook window within Excel). Returns
    True if SetForegroundWindow reported success.
    """
    try:
        import ctypes

        try:
            app.api.WindowState = -4143  # xlNormal (un-minimize)
        except Exception:
            pass
        try:
            app.api.Visible = True
        except Exception:
            pass

        hwnd = int(app.api.Hwnd)
        user32 = ctypes.windll.user32

        SW_RESTORE = 9
        user32.ShowWindow(hwnd, SW_RESTORE)

        # SetForegroundWindow is rejected unless the calling thread "owns" the
        # foreground. Standard trick: attach our input thread to the foreground
        # window's thread, then set foreground, then detach.
        fg = user32.GetForegroundWindow()
        cur_tid = user32.GetWindowThreadProcessId(fg, None)
        our_tid = ctypes.windll.kernel32.GetCurrentThreadId()
        user32.AttachThreadInput(cur_tid, our_tid, True)
        try:
            user32.BringWindowToTop(hwnd)
            ok = bool(user32.SetForegroundWindow(hwnd))
        finally:
            user32.AttachThreadInput(cur_tid, our_tid, False)
        time.sleep(0.5)
        return ok
    except Exception as e:
        logging.warning(f"_force_excel_foreground failed: {e}")
        return False


def trigger_ace_refresh_via_ribbon(app) -> bool:
    """
    Click the ACE 'Refresh' ribbon button by replaying its KeyTip accelerators.
    Excel MUST be the OS-foreground window or the keystrokes go elsewhere
    (e.g. the console). We force foreground via Win32 first.

    Uses the Win32 keybd_event path (not Application.SendKeys) so the keystrokes
    target the actual foreground Excel window reliably.
    Returns True if keystrokes were sent (caller should verify the row count).
    """
    try:
        fg_ok = _force_excel_foreground(app)
        logging.info(f"ACE ribbon refresh: Excel foreground={fg_ok}")
        time.sleep(0.8)

        import ctypes
        user32 = ctypes.windll.user32

        VK = {
            "ALT": 0x12, "Y": 0x59,
            "0": 0x30, "1": 0x31, "2": 0x32, "3": 0x33, "4": 0x34,
            "5": 0x35, "6": 0x36, "7": 0x37, "8": 0x38, "9": 0x39,
            "A": 0x41, "R": 0x52,
        }
        KEYEVENTF_KEYUP = 0x0002

        def tap(vk):
            user32.keybd_event(vk, 0, 0, 0)
            time.sleep(0.05)
            user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
            time.sleep(0.15)

        def send_chars(s):
            for ch in s.upper():
                if ch in VK:
                    tap(VK[ch])
                else:
                    logging.warning(f"  no VK mapping for key '{ch}' — skipping")

        logging.info(
            f"ACE ribbon refresh: Alt -> '{ACE_RIBBON_TAB_KEYS}' -> "
            f"'{ACE_RIBBON_REFRESH_KEY}'"
        )
        # Press & release Alt to surface KeyTips
        tap(VK["ALT"])
        time.sleep(0.6)
        if ACE_RIBBON_TAB_KEYS:
            send_chars(ACE_RIBBON_TAB_KEYS)
            time.sleep(0.9)
        send_chars(ACE_RIBBON_REFRESH_KEY)
        time.sleep(1.0)
        return True
    except Exception as e:
        logging.warning(f"ACE ribbon refresh failed: {e}")
        return False


def ensure_ace_addins(app) -> bool:
    """
    Force-load both ACE add-in components into the given xlwings App.
    Returns True if the UDF-providing .xll appears loaded, else False.
    Safe to call repeatedly; logs but never raises.
    """
    xll_ok  = False
    com_ok  = False

    # --- 1. Load the .xll (the UDF provider) -------------------------------
    try:
        if os.path.exists(ACE_XLL_PATH):
            try:
                # RegisterXLL loads the add-in and registers its functions
                app.api.RegisterXLL(ACE_XLL_PATH)
                xll_ok = True
                logging.info(f"ACE .xll registered: {ACE_XLL_PATH}")
            except Exception as e1:
                # Fallback: add via AddIns2 collection and mark installed
                try:
                    ai = app.api.AddIns2.Add(ACE_XLL_PATH)
                    ai.Installed = True
                    xll_ok = True
                    logging.info(f"ACE .xll installed via AddIns2: {ACE_XLL_PATH}")
                except Exception as e2:
                    logging.warning(f"ACE .xll load failed: RegisterXLL={e1}; AddIns2={e2}")
        else:
            logging.warning(
                f"ACE .xll not found at '{ACE_XLL_PATH}'. "
                f"Set ACE_XLL_PATH env var to the correct path. "
                f"Data cells will be #NAME? until this is fixed."
            )
    except Exception as e:
        logging.warning(f"ACE .xll load error: {e}")

    # --- 2. Connect the COM add-in (refresh / data plumbing) ---------------
    try:
        for ca in app.api.COMAddIns:
            try:
                pid = str(ca.ProgId)
            except Exception:
                pid = ""
            try:
                desc = str(ca.Description)
            except Exception:
                desc = ""
            # ACE is Excel-DNA: ProgId is an opaque 'Dna.<hash>.0' GUID; the
            # human name 'ACEEQ_XL_NXT' lives in .Description. Match on BOTH.
            key = ACE_COM_PROGID.upper()
            if key in pid.upper() or key in desc.upper():
                try:
                    if not ca.Connect:
                        ca.Connect = True
                    com_ok = True
                    logging.info(
                        f"ACE COM add-in connected: ProgId={pid!r} Desc={desc!r}"
                    )
                except Exception as ce:
                    logging.warning(f"ACE COM add-in connect failed ({pid}): {ce}")
                break
        if not com_ok:
            logging.warning(
                f"ACE COM add-in '{ACE_COM_PROGID}' not found in COMAddIns "
                f"collection — company-list refresh may stay stale."
            )
    except Exception as e:
        logging.warning(f"ACE COM add-in enumeration error: {e}")

    return xll_ok


def mark_skipped(company_name, reason):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(SKIPPED_LOG, "a", encoding="utf-8") as f:
        f.write(f"{ts} | {company_name} | {reason}\n")



def safe_float(v):
    if v is None:
        return 0.0
    try:
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).replace(",", "").strip()
        if s in ("-", "--", ""):
            return 0.0
        if s.startswith("(") and s.endswith(")"):
            return -float(s[1:-1])
        return float(s)
    except Exception:
        return 0.0


def update_metadata(company_name):
    try:
        meta_path = os.path.join(DATA_DIR, "metadata.json")
        meta = {}
        if os.path.exists(meta_path):
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
        meta[company_name] = {
            "uploader":  "Automation-Annual",
            "timestamp": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        }
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=4)
        logging.info(f"metadata.json updated for {company_name}")
    except Exception as e:
        logging.warning(f"metadata update failed: {e}")


TRACKER_FILE = os.path.join(DATA_DIR, "annual_tracker.json")
_tracker_lock = Lock()


def _load_tracker() -> dict:
    """Load today's section from annual_tracker.json, creating it if absent."""
    today = datetime.now().strftime("%Y-%m-%d")
    try:
        if os.path.exists(TRACKER_FILE):
            with open(TRACKER_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get(today, {})
    except Exception:
        pass
    return {}


def _save_tracker(today_data: dict):
    today = datetime.now().strftime("%Y-%m-%d")
    try:
        data = {}
        if os.path.exists(TRACKER_FILE):
            with open(TRACKER_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        data[today] = today_data
        with open(TRACKER_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        logging.warning(f"annual_tracker.json save failed: {e}")


def tracker_update(company_name: str, status: str, period: str = "", reason: str = ""):
    """
    Update a company's status in annual_tracker.json.

    Status values used:
      queued     — added to processing queue by ACE poller
      processing — worker picked it up
      uploaded   — successfully processed and uploaded
      retrying   — data not live yet, re-queued
      failed     — error or skipped permanently
    """
    ts = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    with _tracker_lock:
        today_data = _load_tracker()
        entry = today_data.get(company_name, {})
        entry["status"]    = status
        entry["updated"]   = ts
        if period:
            entry["period"] = period
        if reason:
            entry["reason"] = reason
        today_data[company_name] = entry
        _save_tracker(today_data)


def log_status_summary():
    """Log a human-readable summary of uploaded vs pending companies."""
    with _tracker_lock:
        today_data = _load_tracker()

    uploaded   = [n for n, v in today_data.items() if v.get("status") == "uploaded"]
    pending    = [n for n, v in today_data.items() if v.get("status") in ("queued", "processing", "retrying")]
    failed     = [n for n, v in today_data.items() if v.get("status") == "failed"]

    logging.info(
        f"\n{'='*58}\n"
        f"  ANNUAL TRACKER SUMMARY ({datetime.now().strftime('%d-%m-%Y %H:%M')})\n"
        f"  Uploaded  ({len(uploaded)}): {', '.join(uploaded) or 'none'}\n"
        f"  Pending   ({len(pending)}): {', '.join(pending) or 'none'}\n"
        f"  Failed    ({len(failed)}): {', '.join(failed) or 'none'}\n"
        f"{'='*58}"
    )


def _load_done_today() -> set:
    """
    Return company names that should NOT be re-queued.
    Since ACE appends and never removes companies, any company ever
    successfully uploaded across any date should be skipped permanently.
    Reads all dates in annual_tracker.json and collects every 'uploaded' entry.
    """
    result = set()

    try:
        if os.path.exists(TRACKER_FILE):
            with open(TRACKER_FILE, "r", encoding="utf-8") as f:
                all_data = json.load(f)
            for date_key, day_data in all_data.items():
                for company_name, entry in day_data.items():
                    if entry.get("status") == "uploaded":
                        result.add(company_name)
    except Exception as e:
        logging.warning(f"Could not load uploaded companies from tracker: {e}")

    if result:
        logging.info(
            f"Loaded {len(result)} already-uploaded companies "
            f"(all-time, from annual_tracker.json) — will not re-queue: {result}"
        )
    return result


# ================================================================
# Slack
# ================================================================

def _post_slack(webhook_url: str, msg: str):
    if not webhook_url:
        return
    try:
        resp = requests.post(
            webhook_url,
            json={"text": msg},
            headers={"Content-Type": "application/json"},
            timeout=10,
        )
        if resp.status_code == 200:
            logging.info("Slack notification sent")
        else:
            logging.warning(f"Slack post failed ({resp.status_code}): {resp.text}")
    except Exception as e:
        logging.error(f"Slack error: {e}")


def send_slack_annual_processed(company_name: str):
    msg = f"{company_name} annual result updated on DB"
    _post_slack(SLACK_PROCESSED_WEBHOOK, msg)


# ================================================================
# ACE Annual Excel Reader
# ================================================================

def parse_ace_period(period_raw) -> tuple | None:
    """
    Parse ACE period value (e.g. 202512, 202603.0) into
    (src_month, src_yy, tgt_month, tgt_yy).

    Examples:
        202512  -> Dec-24 (src) to Dec-25 (tgt)  => (12, 24, 12, 25)
        202603  -> Mar-25 (src) to Mar-26 (tgt)  => (3,  25, 3,  26)
        202506  -> Jun-24 (src) to Jun-25 (tgt)  => (6,  24, 6,  25)
    """
    if period_raw is None:
        return None
    try:
        period_int = int(float(str(period_raw).strip()))
        year4  = period_int // 100          # e.g. 2025
        month  = period_int % 100           # e.g. 12
        if not (1 <= month <= 12):
            return None
        tgt_yy = year4 % 100               # e.g. 25
        src_yy = (year4 - 1) % 100         # e.g. 24
        return (month, src_yy, month, tgt_yy)
    except Exception:
        return None


def read_ace_annual_companies() -> list:
    """
    Open Annual Update Ace.xlsx in a dedicated Excel instance, call RefreshAll,
    then read company names (col B) and P&L periods (col I) from rows 3+.
    Returns list of dicts: {"company_name": ..., "period": "202512", ...}.
    """
    if not os.path.exists(ANNUAL_ACE_EXCEL):
        logging.error(f"ACE Annual Excel not found: {ANNUAL_ACE_EXCEL}")
        return []

    ace_app = None
    wb      = None
    try:
        logging.info(f"Opening ACE Annual Excel: {ANNUAL_ACE_EXCEL}")
        ace_app = xw.App(visible=True, add_book=False)
        ace_app.display_alerts = False
        try:
            ace_app.api.AutomationSecurity = 1  # msoAutomationSecurityLow
        except Exception:
            pass

        # Load ACE add-ins BEFORE opening/refreshing — automation instances
        # do not auto-load them, which leaves the company list stale.
        ensure_ace_addins(ace_app)

        wb = ace_app.books.open(ANNUAL_ACE_EXCEL)

        # Re-assert after the book is open (some add-ins hook on workbook open)
        ensure_ace_addins(ace_app)

        # Row count BEFORE refresh, to verify the ribbon Refresh actually grew it
        try:
            rows_before = wb.sheets[0].used_range.last_cell.row
        except Exception:
            rows_before = -1
        logging.info(f"ACE rows before refresh: {rows_before}")

        # The ACE company list is refreshed ONLY by the custom 'Refresh' ribbon
        # button (Excel-DNA), NOT by Excel's Data->RefreshAll. Trigger it via
        # the ribbon keyboard accelerator. We still call RefreshAll afterwards
        # for any standard data connections that may exist.
        logging.info("Triggering ACE ribbon Refresh (SendKeys)...")
        trigger_ace_refresh_via_ribbon(ace_app)

        logging.info("Triggering RefreshAll on ACE Annual Excel...")
        try:
            wb.api.RefreshAll()
        except Exception as re_err:
            logging.warning(f"ACE RefreshAll warning (data may be stale): {re_err}")

        logging.info(f"Waiting for ACE data connections to finish (max {ACE_REFRESH_WAIT}s)...")
        elapsed = 0
        while elapsed < ACE_REFRESH_WAIT:
            time.sleep(5)
            elapsed += 5
            try:
                # Check if any workbook connection is still refreshing
                still_refreshing = False
                for conn in wb.api.Connections:
                    try:
                        if conn.OLEDBConnection.Refreshing:
                            still_refreshing = True
                            break
                    except Exception:
                        try:
                            if conn.ODBCConnection.Refreshing:
                                still_refreshing = True
                                break
                        except Exception:
                            pass
                if not still_refreshing:
                    logging.info(f"ACE data connections finished after {elapsed}s")
                    break
            except Exception:
                # If we can't check connections, fall back to CalculationState
                try:
                    if ace_app.api.CalculationState == 0:
                        logging.info(f"ACE Excel done (CalculationState) after {elapsed}s")
                        break
                except Exception:
                    break
            if elapsed % 30 == 0:
                logging.info(f"ACE still refreshing... ({elapsed}s elapsed)")
        else:
            logging.warning(f"ACE Excel did not finish within {ACE_REFRESH_WAIT}s — reading whatever is available")
        # Extra buffer to let the last rows settle after connections report done
        time.sleep(10)

        sheet    = wb.sheets[0]
        last_row = sheet.used_range.last_cell.row
        companies = []

        if rows_before != -1:
            if last_row > rows_before:
                logging.info(
                    f"ACE ribbon Refresh WORKED: rows grew {rows_before} -> {last_row}"
                )
            else:
                logging.warning(
                    f"ACE ribbon Refresh did NOT grow the list "
                    f"({rows_before} -> {last_row}). Check ACE_RIBBON_TAB_KEYS="
                    f"'{ACE_RIBBON_TAB_KEYS}' / ACE_RIBBON_REFRESH_KEY="
                    f"'{ACE_RIBBON_REFRESH_KEY}' and that Excel was frontmost."
                )

        for r in range(3, last_row + 1):            # Row 1=EQNXTQ, Row 2=headers
            name_raw   = sheet.range((r, 2)).value  # Col B = Company Name
            period_raw = sheet.range((r, 9)).value  # Col I = UPDST_Profit and loss Period

            if name_raw is None:
                continue
            name = clean_company_name(str(name_raw).strip())
            if not name:
                continue

            parsed = parse_ace_period(period_raw)
            if parsed is None:
                logging.warning(
                    f"Row {r}: '{name}' has unreadable period '{period_raw}' — skipping"
                )
                continue

            src_month, src_yy, tgt_month, tgt_yy = parsed
            companies.append({
                "company_name": name,
                "period":       str(int(float(str(period_raw).strip()))) if period_raw else "",
                "src_month":    src_month,
                "src_yy":       src_yy,
                "tgt_month":    tgt_month,
                "tgt_yy":       tgt_yy,
            })
            logging.info(
                f"  {name}: period={period_raw} -> "
                f"copy {MONTH_MAP_INV.get(src_month,'?')}-{src_yy:02d} -> "
                f"{MONTH_MAP_INV.get(tgt_month,'?')}-{tgt_yy:02d}"
            )

        logging.info(f"ACE Annual Excel: found {len(companies)} companies with valid periods")
        return companies

    except Exception as e:
        logging.error(f"Failed to read ACE Annual Excel: {e}")
        return []
    finally:
        try:
            if wb:
                wb.close()
        except Exception:
            pass
        try:
            if ace_app:
                ace_app.quit()
        except Exception:
            pass


# ================================================================
# API Download / Upload Wrappers
# ================================================================

def api_download(company_name: str) -> str:
    """
    Download a company Excel file via the API client.

    Checks the Charcha allow-list first; raises CompanyNotFoundError if the
    company is not in the list.

    Returns the local file path of the saved .xlsx in RESULT_FILES_DIR.
    """
    # Safety net: never attempt a download for a company not in the Charcha
    # allow-list (guards the worker / --test-one paths too, not just the poller).
    if not is_in_charcha_list(company_name):
        raise CompanyNotFoundError(
            f"'{company_name}' not in Charcha allow-list — not attempting download."
        )

    logging.info(f"Downloading Excel via API: {company_name}")
    return _api_client.download_file(company_name, RESULT_FILES_DIR)


def api_upload(company_name: str, file_path: str):
    """
    Upload an updated company Excel file via the API client.

    The file should be in UPDATED_DIR.  accord_code is not needed for the
    upload endpoint (it only uses company_name), but we pass a placeholder
    for compatibility with APIClient.upload_file().
    """
    logging.info(f"Uploading via API: {company_name}")
    _api_client.upload_file(file_path, company_name=company_name, accord_code="000000")
    logging.info(f"API upload complete: {company_name}")


# ================================================================
# Excel Processing — Annual
# ================================================================

def process_annual_excel_file(
    file_path: str,
    company_name: str,
    src_month: int,
    src_yy: int,
    tgt_month: int,
    tgt_yy: int,
) -> str | None:
    """
    Open the company template Excel, copy src year formulas → tgt year column in
    the ANNUAL section (left of QTR marker), refresh, validate PAT and Sources of
    Funds, and save.  Returns saved path on success, None if data not yet live.
    """
    src_label = f"{MONTH_MAP_INV.get(src_month,'?')}-{src_yy:02d}"
    tgt_label = f"{MONTH_MAP_INV.get(tgt_month,'?')}-{tgt_yy:02d}"
    logging.info(
        f"process_annual_excel_file: {company_name} | "
        f"copy {src_label} -> {tgt_label}"
    )
    file_name = os.path.basename(file_path)
    logging.info(f"Opening Excel: {file_name} for '{company_name}'")

    force_kill_excel()
    time.sleep(2)

    # IMPORTANT: open via the user's NORMAL Excel (os.startfile) and then attach,
    # exactly like the quarterly processor (result_processor_v3.process_excel_file).
    #
    # Do NOT spawn an isolated instance via xw.App(add_book=False): automation-
    # created Excel instances do not load user add-ins, so the ACE/EQNXTQ data
    # UDFs are unrecognized. RefreshAll/CalculateFull then resolves EVERY ACE
    # formula on the sheet (including untouched historical columns) to #NAME?,
    # and wb.save() persists the corruption. Launching the file normally keeps
    # the ACE add-in loaded so the UDFs resolve.
    os.startfile(file_path)
    # Initial wait for Excel to launch; the attach loop below retries 15x/2s so
    # a slower open is still caught. Env-overridable (EXCEL_OPEN_WAIT).
    time.sleep(int(os.environ.get("EXCEL_OPEN_WAIT", "6")))

    wb = None
    for _ in range(15):
        try:
            for b in xw.books:
                if file_name.lower() in b.name.lower():
                    wb = b
                    break
                stem = os.path.splitext(file_name)[0].lower()
                if stem in b.name.lower():
                    wb = b
                    break
            if wb:
                break
        except Exception:
            pass
        time.sleep(2)

    if not wb:
        logging.error(f"Could not attach to workbook: {file_name}")
        return None

    app = wb.app
    app.display_alerts = False
    # Suppress the external-links "Update links?" prompt on the normal instance
    # (replaces the old update_links=False open arg, which is unavailable when
    # the file is opened via os.startfile).
    try:
        app.api.AskToUpdateLinks = False
    except Exception:
        pass
    try:
        app.api.AutomationSecurity = 1  # msoAutomationSecurityLow
    except Exception:
        pass

    # Ensure the ACE add-ins are loaded in THIS instance before any recalc.
    # Even when opened via os.startfile, assert them so the 'Annual Raw' UDFs
    # resolve and RefreshAll/CalculateFull never wipes the sheet to #NAME?.
    ensure_ace_addins(app)
    time.sleep(2)

    # Pick the right sheet and activate it so Excel's visible state matches
    sheet = None
    for s in wb.sheets:
        if any(n in s.name.upper() for n in ["TEMPLATE", "PRESENTATION"]):
            sheet = s
            break
    if not sheet:
        sheet = wb.sheets[0]
    sheet.activate()   # force Excel to show this sheet (avoids Share price Daily etc.)
    logging.info(f"Using sheet: '{sheet.name}' (activated)")

    # ------------------------------------------------------------------
    # STEP 1 — Scan Column A for section markers
    # ------------------------------------------------------------------
    rows_scan  = sheet.range("A1:A500").value or []
    pl_start   = -1
    val_end    = -1
    pres_start = -1
    pres_end   = -1
    pat_row    = -1
    bs_start   = -1
    bs_end     = -1
    sof_row    = -1

    for r_idx, label in enumerate(rows_scan):
        if not label:
            continue
        lbl = str(label).strip().upper()

        if ("PROFIT & LOSS" in lbl or "P&L" in lbl) and "START" in lbl:
            if pl_start == -1:
                pl_start = r_idx + 1
                logging.info(f"P&L (START) at row {pl_start}")

        if "VALUATION" in lbl and "END" in lbl:
            if val_end == -1:
                val_end = r_idx + 1
                logging.info(f"Valuation (END) at row {val_end}")

        if "PRESENTATION" in lbl and "DATA" in lbl and "START" in lbl:
            if pres_start == -1:
                pres_start = r_idx + 1
                logging.info(f"Presentation Data (Start) at row {pres_start}")

        if "PRESENTATION" in lbl and "DATA" in lbl and "END" in lbl:
            if pres_end == -1:
                pres_end = r_idx + 1
                logging.info(f"Presentation Data (End) at row {pres_end}")

        if "BALANCE SHEET" in lbl and "START" in lbl:
            if bs_start == -1:
                bs_start = r_idx + 1
                logging.info(f"Balance Sheet (Start) at row {bs_start}")

        if "BALANCE SHEET" in lbl and "END" in lbl:
            if bs_end == -1:
                bs_end = r_idx + 1
                logging.info(f"Balance Sheet (End) at row {bs_end}")

    # PAT row (within P&L section)
    if pl_start != -1:
        search_lim = val_end if val_end != -1 else min(pl_start + 300, 500)
        pat_kws    = ["PAT", "PROFIT AFTER TAX", "NET INCOME", "NET PROFIT"]
        for r_idx in range(pl_start - 1, search_lim):
            lbl_raw = rows_scan[r_idx]
            if not lbl_raw:
                continue
            lbl = str(lbl_raw).strip().upper()
            if any(k in lbl for k in pat_kws) and "%" not in lbl:
                if pat_row == -1 or len(lbl) < len(
                    str(rows_scan[pat_row - 1]).strip().upper()
                ):
                    pat_row = r_idx + 1
                    logging.info(f"PAT row at {pat_row}: '{lbl_raw}'")

    # Sources of Funds row (within Balance Sheet section)
    if bs_start != -1:
        bs_search_end = bs_end if bs_end != -1 else min(bs_start + 200, 500)
        sof_kws = ["SOURCES OF FUNDS", "SOURCE OF FUNDS",
                   "TOTAL SOURCES", "TOTAL SOURCE"]
        for r_idx in range(bs_start - 1, bs_search_end):
            lbl_raw = rows_scan[r_idx]
            if not lbl_raw:
                continue
            lbl = str(lbl_raw).strip().upper()
            if any(k in lbl for k in sof_kws):
                sof_row = r_idx + 1
                logging.info(f"Sources of Funds row at {sof_row}: '{lbl_raw}'")
                break

    logging.info(
        f"Markers -> PL_START={pl_start}, VAL_END={val_end}, "
        f"PRES_START={pres_start}, PRES_END={pres_end}, PAT={pat_row}, "
        f"BS_START={bs_start}, BS_END={bs_end}, SOF={sof_row}"
    )

    if pl_start == -1:
        logging.warning(f"Missing P&L START marker for {company_name} — skipping")
        wb.close()
        return None

    # ------------------------------------------------------------------
    # STEP 2 — Read header row 1, find QTR marker to locate annual section
    # ------------------------------------------------------------------
    used_max_col = sheet.used_range.last_cell.column
    last_row     = sheet.used_range.last_cell.row
    header       = sheet.range((1, 1), (1, used_max_col)).value or []

    copy_end = last_row
    logging.info(f"Sheet last used row: {last_row}")

    # Annual columns sit LEFT of the QTR/QUARTERLY divider.
    # Find QTR marker to define the right boundary of the annual search.
    qtr_col = -1
    for c_idx, val in enumerate(header):
        if val and ("QTR" in str(val).upper().replace(" ", "")
                    or "QUARTERLY" in str(val).upper().replace(" ", "")):
            qtr_col = c_idx   # 0-based index
            logging.info(f"QTR marker at col {c_idx + 1}: '{val}' — annual search boundary")
            break

    # Search range: full header up to (but not including) the QTR marker.
    # If no QTR marker found, search the full header.
    search_end = qtr_col if qtr_col != -1 else len(header)
    logging.info(
        f"Annual column search range: cols 1-{search_end} "
        f"({'before QTR marker' if qtr_col != -1 else 'full header, no QTR marker'})"
    )

    # ------------------------------------------------------------------
    # STEP 3 — Find src and tgt columns in annual section (before QTR marker)
    # ------------------------------------------------------------------
    src_col = -1
    tgt_col = -1

    for c_idx in range(search_end):
        val = header[c_idx]
        col = c_idx + 1

        if src_col == -1 and date_matches(val, src_month, src_yy):
            src_col = col
            logging.info(f"{src_label} (source) found at col {col}: '{val}'")

        if tgt_col == -1 and date_matches(val, tgt_month, tgt_yy):
            tgt_col = col
            logging.info(f"{tgt_label} (target) found at col {col}: '{val}'")

        if src_col != -1 and tgt_col != -1:
            break

    if src_col == -1:
        # ACE-derived source month not found — fall back to the two most recent
        # date-header columns in the annual section, regardless of month.
        # This handles companies that changed fiscal year (e.g. June → March).
        dated_cols = []   # list of (col_1based, parsed_date_tuple)
        for c_idx in range(search_end):
            parsed = parse_header_date(header[c_idx])
            if parsed:
                dated_cols.append((c_idx + 1, parsed))

        if len(dated_cols) >= 2:
            # Sort by (year, month) — newest last
            dated_cols.sort(key=lambda x: (x[1][1], x[1][0]))  # sort by (yy, month)
            fb_src_col,  fb_src_date  = dated_cols[-2]
            fb_tgt_col,  fb_tgt_date  = dated_cols[-1]
            fb_src_label = f"{MONTH_MAP_INV.get(fb_src_date[0],'?')}-{fb_src_date[1]:02d}"
            fb_tgt_label = f"{MONTH_MAP_INV.get(fb_tgt_date[0],'?')}-{fb_tgt_date[1]:02d}"
            logging.warning(
                f"{company_name}: ACE period {src_label}->{tgt_label} not found in "
                f"annual section (fiscal year mismatch?). "
                f"Falling back to last two annual columns: "
                f"col {fb_src_col} ({fb_src_label}) -> col {fb_tgt_col} ({fb_tgt_label})"
            )
            src_col   = fb_src_col
            tgt_col   = fb_tgt_col
            src_label = fb_src_label + " [fallback]"
            tgt_label = fb_tgt_label + " [fallback]"
        else:
            ann_headers = [
                f"col{i+1}={header[i]!r}"
                for i in range(search_end)
                if header[i] is not None
            ]
            logging.warning(
                f"{company_name}: {src_label} not found and not enough dated "
                f"columns for fallback (found {len(dated_cols)}). "
                f"Annual headers: {ann_headers}"
            )
            wb.close()
            return None

    if tgt_col == -1:
        # Target column doesn't exist yet — create it one position after source
        tgt_col = src_col + 1
        logging.info(f"{tgt_label} not in header -> will be created at col {tgt_col}")

    logging.info(
        f"COPY: col {src_col} ({src_label}) -> col {tgt_col} ({tgt_label})  "
        f"[annual section, QTR boundary at col "
        f"{qtr_col + 1 if qtr_col != -1 else 'N/A (no QTR marker)'}]"
    )

    # Alias to existing variable names so the steps below still work unchanged
    mar25_col = src_col
    mar26_col = tgt_col

    # ------------------------------------------------------------------
    # STEP 4 — Build row ranges (exclude Presentation Data)
    # ------------------------------------------------------------------
    copy_ranges = []
    if pres_start != -1 and pres_end != -1:
        if pl_start <= pres_start - 1:
            copy_ranges.append((pl_start, pres_start - 1))
        if pres_end + 1 <= copy_end:
            copy_ranges.append((pres_end + 1, copy_end))
    else:
        copy_ranges.append((pl_start, copy_end))

    logging.info(f"Row ranges to copy (excl Pres Data): {copy_ranges}")

    # ------------------------------------------------------------------
    # STEP 5 — Write MAR-26 header if column is new
    # ------------------------------------------------------------------
    if mar26_col > len(header):
        import calendar as _cal
        tgt_year4 = 2000 + tgt_yy
        tgt_day   = _cal.monthrange(tgt_year4, tgt_month)[1]
        mar25_raw = header[mar25_col - 1]
        if isinstance(mar25_raw, (datetime, pd.Timestamp)):
            mar26_header_val = datetime(tgt_year4, tgt_month, tgt_day)
        else:
            s = str(mar25_raw).strip()
            mar26_header_val = re.sub(str(src_yy).zfill(2), str(tgt_yy).zfill(2), s)

        sheet.range((1, mar26_col)).value = mar26_header_val
        logging.info(f"Wrote {tgt_label} header at col {mar26_col}: '{mar26_header_val}'")
    else:
        logging.info(f"{tgt_label} header already exists at col {mar26_col}")

    # ------------------------------------------------------------------
    # STEP 5b — Backup existing MAR-26 → Historical Estimates column
    # ------------------------------------------------------------------
    try:
        he_header = sheet.range((1, 1), (1, used_max_col + 10)).value or []

        hist_est_col = -1
        for c_idx, val in enumerate(he_header):
            if val and "HISTORICAL" in str(val).strip().upper() \
                    and "ESTIMATE" in str(val).strip().upper():
                hist_est_col = c_idx + 1
                logging.info(
                    f"Found 'Historical Estimates' marker at col {hist_est_col}: '{val}'"
                )
                break

        if hist_est_col != -1:
            backup_col = hist_est_col + 1
            logging.info(f"Will backup MAR-26 values to col {backup_col}")
        else:
            current_last_col = sheet.used_range.last_cell.column
            hist_est_col = current_last_col + 4
            backup_col   = hist_est_col + 1
            sheet.range((1, hist_est_col)).value = "HISTORICAL ESTIMATES"
            logging.info(
                f"Created 'HISTORICAL ESTIMATES' header at col {hist_est_col}"
            )

        # Set the backup column header only if it's empty (first-write-wins,
        # consistent with the data policy below).
        existing_backup_header = (
            he_header[backup_col - 1] if backup_col <= len(he_header) else None
        )
        if existing_backup_header is None or str(existing_backup_header).strip() == "":
            if mar26_col <= len(he_header) and he_header[mar26_col - 1]:
                backup_header_val = he_header[mar26_col - 1]
            else:
                backup_header_val = "Mar-26"
            sheet.range((1, backup_col)).value = backup_header_val

        # FIRST-WRITE-WINS BACKUP (never overwrite existing backup data).
        # The Historical Estimates backup must capture the ORIGINAL Mar-26
        # estimates exactly once. On any RE-RUN of a company, the downloaded
        # file's Mar-26 may be corrupted (#NAME?) or stale/wrong numbers — and
        # copying that over an already-populated backup would destroy the only
        # good copy of the estimates. So we write a backup cell ONLY when it is
        # currently EMPTY; any row that already has a backup value is left
        # untouched, regardless of what the source now holds.
        src_vals = sheet.range((pl_start, mar26_col),
                               (copy_end, mar26_col)).value
        bak_vals = sheet.range((pl_start, backup_col),
                               (copy_end, backup_col)).value
        # Normalize to flat lists (single-row ranges come back as scalars)
        if not isinstance(src_vals, list):
            src_vals = [src_vals]
        if not isinstance(bak_vals, list):
            bak_vals = [bak_vals]

        def _is_empty(v):
            return v is None or (isinstance(v, str) and v.strip() == "")

        out_vals = []
        written  = 0
        kept     = 0
        for i in range(len(src_vals)):
            s = src_vals[i]
            b = bak_vals[i] if i < len(bak_vals) else None
            if _is_empty(b):
                # Backup cell is empty -> first write: capture the source value
                out_vals.append(s)
                written += 1
            else:
                # Backup already populated -> NEVER overwrite (first-write-wins)
                out_vals.append(b)
                kept += 1

        # Write the merged column back (values only, as a column vector)
        dst_backup = sheet.range((pl_start, backup_col), (copy_end, backup_col))
        dst_backup.value = [[v] for v in out_vals]

        logging.info(
            f"Historical Estimates backup (col {backup_col}), rows "
            f"{pl_start}->{copy_end}: {written} cells newly captured, "
            f"{kept} preserved (already had a backup value — NOT overwritten)"
        )

    except Exception as he_err:
        logging.warning(
            f"Historical Estimates backup failed (non-fatal): {he_err}. "
            f"Proceeding with formula copy anyway."
        )

    time.sleep(1)

    # ------------------------------------------------------------------
    # STEP 6 — Copy MAR-25 formulas → MAR-26
    # ------------------------------------------------------------------
    for seg_start, seg_end in copy_ranges:
        try:
            src = sheet.range((seg_start, mar25_col), (seg_end, mar25_col))
            dst = sheet.range((seg_start, mar26_col), (seg_end, mar26_col))
            src.api.Copy(Destination=dst.api)
            logging.info(
                f"Copied MAR-25 -> MAR-26: rows {seg_start}->{seg_end}, "
                f"col {mar25_col} -> col {mar26_col}"
            )
        except Exception as e:
            logging.warning(f"Copy-paste failed for rows {seg_start}-{seg_end}: {e}")

    try:
        app.api.CutCopyMode = False
    except Exception:
        pass

    time.sleep(2)

    # ------------------------------------------------------------------
    # STEP 7 — Pre-Refresh PAT
    # ------------------------------------------------------------------
    pre_pat = None
    if pat_row != -1 and mar26_col != -1:
        pre_pat = sheet.range((pat_row, mar26_col)).value
        logging.info(f"PRE-refresh PAT at ({pat_row},{mar26_col}): {pre_pat}")

    # ------------------------------------------------------------------
    # STEP 8 — Refresh ACE data
    # ------------------------------------------------------------------
    # The newly-copied target column pulls its values from the ACE 'Annual Raw'
    # sheet, which is populated by the ACE add-in. As proven on the master
    # company list, Excel's wb.RefreshAll() does NOT trigger an ACE data pull —
    # only the custom ACE ribbon 'Refresh' button does (Alt -> Y2 -> Y8 via
    # Win32 foreground + keybd_event). So trigger the ribbon Refresh FIRST, then
    # also call RefreshAll/CalculateFull for any standard connections + to
    # recompute the formula references.
    time.sleep(3)
    logging.info(f"Triggering ACE ribbon Refresh for {company_name} (SendKeys)...")
    trigger_ace_refresh_via_ribbon(app)

    logging.info(f"Triggering RefreshAll/CalculateFull for {company_name}...")
    try:
        wb.api.RefreshAll()
        app.api.CalculateFull()
    except Exception as ex:
        logging.warning(f"Refresh error: {ex}")

    logging.info(f"Waiting {EXCEL_REFRESH_WAIT}s for data pull...")
    time.sleep(EXCEL_REFRESH_WAIT)

    try:
        sheet = wb.sheets[sheet.name]
    except Exception:
        pass

    # ------------------------------------------------------------------
    # STEP 9 — Post-Refresh PAT
    # ------------------------------------------------------------------
    post_pat = None
    if pat_row != -1 and mar26_col != -1:
        try:
            post_pat = sheet.range((pat_row, mar26_col)).value
        except Exception:
            sheet    = wb.sheets[sheet.name]
            post_pat = sheet.range((pat_row, mar26_col)).value
    logging.info(f"POST-refresh PAT: {post_pat}")

    v_pre  = safe_float(pre_pat)
    v_post = safe_float(post_pat)
    diff   = abs(v_post - v_pre)
    logging.info(
        f"PAT delta for {company_name}: {diff:.6f}  "
        f"(pre={v_pre:.4f}, post={v_post:.4f})"
    )

    # ------------------------------------------------------------------
    # STEP 9b — Post-Refresh Sources of Funds
    # ------------------------------------------------------------------
    sof_value = None
    sof_ok    = True

    if sof_row != -1 and mar26_col != -1:
        try:
            sof_value = sheet.range((sof_row, mar26_col)).value
        except Exception:
            try:
                sheet     = wb.sheets[sheet.name]
                sof_value = sheet.range((sof_row, mar26_col)).value
            except Exception:
                pass

        sof_float = safe_float(sof_value)
        sof_ok    = abs(sof_float) > 0.0001
        logging.info(
            f"Sources of Funds at ({sof_row},{mar26_col}): "
            f"{sof_value}  (float={sof_float:.4f}, ok={sof_ok})"
        )
    else:
        logging.info("Sources of Funds row not found — skipping BS check")

    # ------------------------------------------------------------------
    # STEP 9c — Corruption guard: refuse to save if the target column is
    # full of #NAME? errors (ACE add-in failed to load / resolve). Saving
    # here is what pushed malformed files to Charcha. Treat as "not live"
    # so the worker re-queues instead of overwriting good data with errors.
    # ------------------------------------------------------------------
    try:
        check_vals = sheet.range((pl_start, mar26_col),
                                 (copy_end, mar26_col)).value or []
        flat = [v for v in (check_vals if isinstance(check_vals, list) else [check_vals])]
        name_errors = sum(1 for v in flat if isinstance(v, str) and "#NAME?" in v)
        if name_errors > 0:
            logging.error(
                f"ABORTING SAVE for {company_name}: target column has "
                f"{name_errors} #NAME? cells — ACE add-in did not resolve. "
                f"File NOT saved (would corrupt Charcha). Re-queuing."
            )
            wb.close()
            return None
    except Exception as guard_err:
        logging.warning(f"#NAME? guard check failed (non-fatal): {guard_err}")

    # ------------------------------------------------------------------
    # STEP 10 — Save (ACE presence confirms data is live)
    # ------------------------------------------------------------------
    logging.info(
        f"Saving {company_name} "
        f"(PAT: pre={v_pre:.4f}, post={v_post:.4f}, SOF={safe_float(sof_value):.4f})"
    )
    final_path = os.path.join(UPDATED_DIR, file_name)
    wb.save(final_path)
    update_metadata(company_name)
    time.sleep(2)
    wb.close()
    logging.info(f"Saved: {final_path}")
    return final_path


# ================================================================
# Worker Thread
# ================================================================

def annual_worker_thread(processing_queue: PriorityQueue):
    logging.info("Annual Worker thread started")

    while not stop_event.is_set():
        try:
            try:
                process_at, _id, item = processing_queue.get(timeout=5)
            except Empty:
                continue

            now = datetime.now()
            if process_at > now:
                wait_secs = (process_at - now).total_seconds()
                logging.info(
                    f"Next: '{item['company_name']}' at "
                    f"{process_at.strftime('%H:%M:%S')} "
                    f"({wait_secs:.0f}s away)"
                )
                processing_queue.put((process_at, _id, item))
                stop_event.wait(min(wait_secs, 30))
                continue

            company_name = item["company_name"]
            attempt      = item.get("attempt", 1)
            src_month    = item.get("src_month", 3)
            src_yy       = item.get("src_yy",    25)
            tgt_month    = item.get("tgt_month", 3)
            tgt_yy       = item.get("tgt_yy",    26)

            logging.info(f"\n{'='*58}")
            logging.info(
                f"[{datetime.now().strftime('%H:%M:%S')}] Annual Processing: "
                f"{company_name}  (attempt #{attempt})"
            )
            logging.info(f"{'='*58}")
            tracker_update(company_name, "processing")

            downloaded = None

            try:
                force_kill_excel()

                try:
                    downloaded = api_download(company_name)
                except CompanyNotFoundError as nf_err:
                    # Name not on Charcha allow-list -> skip FAST, mark failed.
                    logging.warning(f"Company not found '{company_name}': {nf_err}")
                    mark_skipped(company_name, f"Not found: {nf_err}")
                    tracker_update(company_name, "failed", reason="Company not found on Charcha")
                    with state_lock:
                        in_queue.discard(company_name)
                    continue
                except Exception as dl_err:
                    logging.warning(f"API download failed '{company_name}': {dl_err}")
                    mark_skipped(company_name, f"Download: {dl_err}")
                    tracker_update(company_name, "failed", reason=f"Download error: {dl_err}")
                    with state_lock:
                        in_queue.discard(company_name)
                    continue

                try:
                    result_file = process_annual_excel_file(
                        downloaded, company_name,
                        src_month, src_yy, tgt_month, tgt_yy,
                    )
                except Exception as ex_err:
                    logging.error(f"Excel error '{company_name}': {ex_err}")
                    mark_skipped(company_name, f"Excel: {ex_err}")
                    tracker_update(company_name, "failed", reason=f"Excel error: {ex_err}")
                    force_kill_excel()
                    _safe_delete(downloaded)
                    with state_lock:
                        in_queue.discard(company_name)
                    continue

                if result_file and os.path.exists(result_file):
                    try:
                        force_kill_excel()
                        api_upload(company_name, result_file)
                        with state_lock:
                            done_today.add(company_name)
                            in_queue.discard(company_name)
                        logging.info(f"ANNUAL FULLY DONE: {company_name}")
                        tracker_update(company_name, "uploaded")
                        send_slack_annual_processed(company_name)
                    except Exception as ul_err:
                        logging.error(f"Upload failed '{company_name}': {ul_err}")
                        mark_skipped(company_name, f"Upload: {ul_err}")
                        tracker_update(company_name, "failed", reason=f"Upload error: {ul_err}")
                        with state_lock:
                            in_queue.discard(company_name)

                    _safe_delete(downloaded)

                else:
                    logging.info(
                        f"{company_name}: Annual data not yet live -> delete file, "
                        f"retry in {ANNUAL_RETRY_DELAY_MINS} min"
                    )
                    force_kill_excel()
                    _safe_delete(downloaded)

                    retry_at = datetime.now() + timedelta(minutes=ANNUAL_RETRY_DELAY_MINS)
                    item["attempt"] = attempt + 1
                    processing_queue.put((retry_at, id(item), item))
                    tracker_update(
                        company_name, "retrying",
                        reason=f"Data not live yet, retry at {retry_at.strftime('%H:%M:%S')} (attempt #{attempt+1})"
                    )
                    logging.info(
                        f"Re-queued '{company_name}' -> "
                        f"{retry_at.strftime('%H:%M:%S')} (attempt #{attempt+1})"
                    )

            except Exception as fatal:
                logging.error(f"Fatal error for '{company_name}': {fatal}")
                force_kill_excel()
                _safe_delete(downloaded)
                with state_lock:
                    in_queue.discard(company_name)

            force_kill_excel()

        except Exception as outer_err:
            logging.error(f"Worker outer error: {outer_err}")
            time.sleep(5)

    logging.info("Annual Worker thread stopped")


def _safe_delete(path):
    if path and os.path.exists(path):
        try:
            os.remove(path)
            logging.info(f"Deleted: {path}")
        except Exception:
            pass


# ================================================================
# ACE Annual Poller Thread
# ================================================================

def ace_annual_poller_thread(processing_queue: PriorityQueue):
    logging.info("ACE Annual Poller thread started")

    current_date = None

    while not stop_event.is_set():
        today_str = datetime.now().strftime("%d-%m-%Y")

        # Daily reset at midnight (or first run when current_date is None)
        if today_str != current_date:
            logging.info(f"ACE Poller: Date changed to {today_str}. Resetting state.")
            current_date = today_str
            with state_lock:
                done_today.clear()
                in_queue.clear()
                # Only uploaded companies are permanently skipped (all-time, all dates).
                # failed / queued / processing / retrying will be re-queued naturally
                # when ACE Excel is read on the next poll cycle.
                done_today.update(_load_done_today())

        # Read ACE Excel
        try:
            companies = read_ace_annual_companies()
        except Exception as e:
            logging.error(f"ACE poller read error: {e}")
            stop_event.wait(ANNUAL_POLL_INTERVAL)
            continue

        if not companies:
            logging.info("ACE Annual Excel returned 0 companies — will retry next poll")
            stop_event.wait(ANNUAL_POLL_INTERVAL)
            continue

        new_count = 0
        for co in companies:
            company_name = co["company_name"]

            # Skip companies that are not in the Charcha DB allow-list — no point
            # downloading; mark failed and move on.
            if not is_in_charcha_list(company_name):
                logging.info(
                    f"ACE Poller: '{company_name}' not in Charcha allow-list "
                    f"— skipping (marked failed)."
                )
                tracker_update(company_name, "failed",
                               reason="Not in Charcha companies list")
                continue

            with state_lock:
                if company_name in done_today:
                    # Ensure today's tracker reflects this as uploaded (fixes summary display)
                    tracker_update(company_name, "uploaded")
                    continue
                if company_name in in_queue:
                    continue
                in_queue.add(company_name)

            process_at = datetime.now() + timedelta(minutes=ANNUAL_INITIAL_DELAY_MINS)
            item = {
                "company_name": company_name,
                "src_month":    co["src_month"],
                "src_yy":       co["src_yy"],
                "tgt_month":    co["tgt_month"],
                "tgt_yy":       co["tgt_yy"],
                "attempt":      1,
            }
            processing_queue.put((process_at, id(item), item))
            src_lbl = f"{MONTH_MAP_INV.get(co['src_month'],'?')}-{co['src_yy']:02d}"
            tgt_lbl = f"{MONTH_MAP_INV.get(co['tgt_month'],'?')}-{co['tgt_yy']:02d}"
            logging.info(
                f"ACE Poller: Queued '{company_name}' "
                f"({src_lbl} -> {tgt_lbl}) -> "
                f"{process_at.strftime('%H:%M:%S')}"
            )
            tracker_update(company_name, "queued", period=co.get("period", ""))
            new_count += 1

        if new_count:
            logging.info(f"ACE Poller: Added {new_count} new companies to queue")
        else:
            logging.info("ACE Poller: No new companies (all already done or queued)")

        stop_event.wait(ANNUAL_POLL_INTERVAL)

    logging.info("ACE Annual Poller thread stopped")


# ================================================================
# Main
# ================================================================

def run():
    global _api_client

    logging.info("=" * 60)
    logging.info("Annual Result Processor v2 starting (API mode, 24×7)...")
    logging.info(f"EXE directory: {EXE_DIR}")
    logging.info(f"Base directory: {BASE_DIR}")
    logging.info(f"ACE Excel: {ANNUAL_ACE_EXCEL}")
    logging.info(f"Log file: {LOG_FILE}")
    logging.info("=" * 60)

    trust_download_folder(RESULT_FILES_DIR)

    # Initialise the API client (handles token generation/refresh internally)
    _api_logger = logging.getLogger("api_client")
    _api_client = APIClient(_api_logger, _api_logger)
    try:
        _api_client.login()
    except Exception as e:
        logging.warning(f"Initial API login error: {e} — will retry on first use")

    # Pre-populate done_today from metadata.json (survives restart)
    with state_lock:
        done_today.update(_load_done_today())

    # Shared queue
    processing_queue = PriorityQueue()

    # Start threads
    poller_t = Thread(
        target=ace_annual_poller_thread,
        args=(processing_queue,),
        daemon=True,
        name="ACE-Annual-Poller",
    )
    worker_t = Thread(
        target=annual_worker_thread,
        args=(processing_queue,),
        daemon=True,
        name="Annual-Excel-Worker",
    )

    poller_t.start()
    worker_t.start()

    logging.info(
        f"All threads running (ACE Poller every {ANNUAL_POLL_INTERVAL}s + Worker). "
        f"Press Ctrl+C to stop."
    )

    try:
        while True:
            time.sleep(60)
            with state_lock:
                logging.info(
                    f"Status — Queue: ~{processing_queue.qsize()} | "
                    f"Done: {len(done_today)} | Pending: {len(in_queue)}"
                )
            log_status_summary()
    except KeyboardInterrupt:
        logging.info("\nShutting down...")
        stop_event.set()
        poller_t.join(timeout=15)
        worker_t.join(timeout=15)
        if _api_client:
            _api_client.close()
        force_kill_excel()
        logging.info("Shutdown complete.")


def run_forever():
    """
    24×7 resilience wrapper around run().

    If run() raises an unhandled exception, log the full traceback, wait
    RUN_LOOP_COOLDOWN seconds, and restart — unless the stop_event has been
    set (clean shutdown via Ctrl+C).
    """
    while not stop_event.is_set():
        try:
            run()
            break  # run() exited cleanly (e.g. KeyboardInterrupt handled inside)
        except KeyboardInterrupt:
            logging.info("KeyboardInterrupt caught in run_forever — shutting down.")
            stop_event.set()
            break
        except Exception:
            logging.error(
                f"UNHANDLED ERROR in run() — will restart after "
                f"{RUN_LOOP_COOLDOWN}s cooldown.\n{traceback.format_exc()}"
            )
            if stop_event.is_set():
                break
            stop_event.wait(RUN_LOOP_COOLDOWN)
            if stop_event.is_set():
                break
            logging.info("Restarting run() loop...")
            # Reset thread-safe state for fresh start
            with state_lock:
                in_queue.clear()


def diagnose_ace():
    """
    Inspect the ACE add-in setup in an xlwings automation instance and print
    everything needed to programmatically trigger the ACE 'Refresh' ribbon
    command. Invoked via:  python annual_processor.py --diagnose
    Nothing is saved or modified.
    """
    print("=" * 70)
    print(">>> ACE ADD-IN DIAGNOSTIC <<<")
    print("=" * 70)
    print("Opening Excel via xlwings (automation instance)...")
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False

    try:
        app.api.RegisterXLL(ACE_XLL_PATH)
        print(f"RegisterXLL OK: {ACE_XLL_PATH}")
    except Exception as e:
        print(f"RegisterXLL FAILED: {e}")

    print("\n--- COMAddIns collection ---")
    try:
        n = app.api.COMAddIns.Count
        print(f"Count = {n}")
        for i in range(1, n + 1):
            ca = app.api.COMAddIns.Item(i)
            try:
                print(f"  [{i}] ProgId={ca.ProgId!r}  Connect={ca.Connect}  "
                      f"Desc={getattr(ca, 'Description', '?')!r}")
            except Exception as e:
                print(f"  [{i}] (error reading: {e})")
    except Exception as e:
        print(f"COMAddIns enumeration error: {e}")

    print("\n--- AddIns2 collection (.xll/.xlam) ---")
    try:
        for ai in app.api.AddIns2:
            try:
                print(f"  Name={ai.Name!r}  Installed={ai.Installed}  "
                      f"Path={getattr(ai, 'FullName', '?')!r}")
            except Exception as e:
                print(f"  (error: {e})")
    except Exception as e:
        print(f"AddIns2 enumeration error: {e}")

    print("\n--- ACE COM add-in .Object members (if found) ---")
    try:
        for i in range(1, app.api.COMAddIns.Count + 1):
            ca = app.api.COMAddIns.Item(i)
            pid = str(ca.ProgId)
            if any(k in pid.upper() for k in ("ACE", "ACEEQ", "ACCORD")):
                print(f"  Found ACE COM add-in: {pid}")
                try:
                    if not ca.Connect:
                        ca.Connect = True
                        print("    -> set Connect=True")
                except Exception as e:
                    print(f"    -> connect failed: {e}")
                try:
                    obj = ca.Object
                    print(f"    .Object = {obj}")
                    print(f"    dir(.Object) = "
                          f"{[m for m in dir(obj) if not m.startswith('_')]}")
                except Exception as e:
                    print(f"    .Object error: {e}")
    except Exception as e:
        print(f"ACE .Object inspection error: {e}")

    print("\n--- ExecuteMso candidate checks ---")
    for mso in ["RefreshAll", "DataRefreshAll", "Refresh", "QueryRefresh"]:
        try:
            ok = app.api.CommandBars.GetEnabledMso(mso)
            print(f"  {mso}: GetEnabledMso={ok}")
        except Exception as e:
            print(f"  {mso}: invalid/err ({e})")

    # --- Inspect the ACE install dir for the Excel-DNA .dna / ribbon XML ---
    # Excel-DNA stores the ribbon definition (and the real onAction callback
    # name for the Refresh button) in a .dna file next to the .xll, or packed
    # inside the .xll. Find and print any .dna/.config/.xml so we get the
    # EXACT callback name instead of guessing.
    discovered = []   # onAction callback names recovered from .dna / .xll
    print("\n--- ACE install-dir scan (looking for .dna / ribbon XML) ---")
    try:
        import re as _re
        ace_dir = os.path.dirname(ACE_XLL_PATH)
        print(f"  dir: {ace_dir}")
        for fn in os.listdir(ace_dir):
            low = fn.lower()
            full = os.path.join(ace_dir, fn)
            try:
                print(f"    {fn}  ({os.path.getsize(full)} bytes)")
            except Exception:
                print(f"    {fn}")
            if low.endswith((".dna", ".config", ".xml")):
                try:
                    with open(full, "r", encoding="utf-8", errors="replace") as f:
                        txt = f.read()
                    hits = _re.findall(r'.{0,40}(?:onAction|customUI|getLabel|'
                                       r'[Rr]efresh|idMso|id=).{0,60}', txt)
                    print(f"      --- ribbon-relevant lines in {fn} ---")
                    for h in hits[:60]:
                        print(f"        {h.strip()}")
                    for m in _re.findall(r'onAction=["\']([^"\']{1,60})["\']', txt):
                        discovered.append(m)
                except Exception as e:
                    print(f"      (could not read {fn}: {e})")
    except Exception as e:
        print(f"  install-dir scan error: {e}")

    # --- Search the .xll binary for embedded onAction callback names ---
    # Excel-DNA often packs the ribbon XML inside the .xll. Grep ASCII strings
    # for onAction="..." patterns to recover the Refresh callback name.
    print("\n--- .xll embedded onAction scan ---")
    try:
        import re as _re
        with open(ACE_XLL_PATH, "rb") as f:
            blob = f.read()
        text = blob.decode("latin-1", errors="replace")
        for pat in [r'onAction="([^"]{1,60})"',
                    r'onAction=\x27([^\x27]{1,60})\x27']:
            for m in set(_re.findall(pat, text)):
                print(f"    onAction callback: {m!r}")
                discovered.append(m)
        # Also any 'customUI' marker confirming embedded ribbon
        if "customUI" in text:
            print("    (.xll contains embedded customUI ribbon XML)")
    except Exception as e:
        print(f"  .xll scan error: {e}")

    # --- Open the ACE workbook and try to trigger the refresh via macro ---
    # Excel-DNA add-ins expose ribbon callbacks as macro-callable functions.
    # We open the real workbook, count rows, try each candidate macro name via
    # Application.Run, wait, and re-count. A name that grows the count is THE fix.
    print("\n--- Application.Run refresh-macro probe ---")
    wb = None
    try:
        if not os.path.exists(ANNUAL_ACE_EXCEL):
            print(f"  ACE workbook not found: {ANNUAL_ACE_EXCEL} — skipping macro probe")
        else:
            wb = app.books.open(ANNUAL_ACE_EXCEL)
            sht = wb.sheets[0]

            def row_count():
                try:
                    return sht.used_range.last_cell.row
                except Exception:
                    return -1

            before = row_count()
            print(f"  Workbook opened. used rows BEFORE = {before}")

            # Discovered onAction names first, then fallback guesses.
            candidates = list(dict.fromkeys(discovered + [
                "Refresh", "RefreshAll", "RefreshAllSheets", "Refresh_All_Sheets",
                "ACEEQ_XL_NXT.Refresh", "ACEEQ_Refresh", "AceRefresh",
                "RibbonRefresh", "OnRefresh", "btnRefresh_Click",
                "RefreshData", "RefreshSheet", "UpdateData",
            ]))
            print(f"  Discovered onAction names: {discovered or '(none)'}")
            for name in candidates:
                try:
                    app.api.Run(name)
                    print(f"  Application.Run('{name}') -> OK (no error)")
                    time.sleep(8)
                    after = row_count()
                    print(f"       used rows AFTER = {after}"
                          f"{'   *** GREW! THIS IS THE MACRO ***' if after > before else ''}")
                    if after > before:
                        break
                except Exception as e:
                    msg = str(e)
                    short = msg[:90]
                    print(f"  Application.Run('{name}') -> ERR {short}")
    except Exception as e:
        print(f"  macro probe error: {e}")

    print("\nDone. Closing without saving, quitting Excel.")
    try:
        if wb:
            wb.close()  # no save
    except Exception:
        pass
    try:
        app.quit()
    except Exception:
        pass


def test_refresh():
    """
    Fast isolated test of the ACE ribbon Refresh (no Charcha, no queue).
    Opens the ACE workbook, triggers the ribbon Refresh via SendKeys, waits,
    and reports whether the row count grew. Closes WITHOUT saving.
    Invoked via:  python annual_processor.py --test-refresh
    """
    print("=" * 60)
    print(">>> ACE RIBBON REFRESH TEST <<<")
    print(f"  TAB_KEYS='{ACE_RIBBON_TAB_KEYS}'  REFRESH_KEY='{ACE_RIBBON_REFRESH_KEY}'")
    print("=" * 60)
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    wb = None
    try:
        ensure_ace_addins(app)
        wb = app.books.open(ANNUAL_ACE_EXCEL)
        ensure_ace_addins(app)
        before = wb.sheets[0].used_range.last_cell.row
        print(f"  rows BEFORE = {before}")
        trigger_ace_refresh_via_ribbon(app)
        print("  waiting 60s for refresh to populate...")
        time.sleep(60)
        after = wb.sheets[0].used_range.last_cell.row
        print(f"  rows AFTER  = {after}")
        print("  *** REFRESH WORKED ***" if after > before
              else "  XXX refresh did NOT grow the list XXX")
    except Exception as e:
        print(f"  error: {e}")
    finally:
        try:
            if wb:
                wb.close()
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            pass


def test_one(company_name: str, period: str = ""):
    """
    Process exactly ONE company end-to-end (download -> fill -> ACE refresh ->
    save) WITHOUT uploading, so the saved Updated_Excel file can be inspected
    for #NAME? / correct numbers.

    Does NOT open/refresh the master Annual Update Ace.xlsx (that step is slow
    and irrelevant here). The period defaults to the standard Mar-25 -> Mar-26;
    override by passing an ACE period as the 2nd arg, e.g.:
        python annual_processor.py --test-one "Castrol India Ltd" 202512
    Invoked via:  python annual_processor.py --test-one "Company Name Ltd"
    """
    global _api_client

    print("=" * 60)
    print(f">>> TEST ONE COMPANY (no upload): {company_name} <<<")
    print("=" * 60)

    # Resolve the period WITHOUT touching the master file.
    parsed = parse_ace_period(period) if period else None
    if parsed is None:
        if period:
            print(f"  period '{period}' unparseable — defaulting to 202603 (Mar-25 -> Mar-26)")
        src_month, src_yy, tgt_month, tgt_yy = (3, 25, 3, 26)
    else:
        src_month, src_yy, tgt_month, tgt_yy = parsed
    print(f"  using period -> copy {src_month}/{src_yy:02d} -> {tgt_month}/{tgt_yy:02d}")

    # Initialise API client for download
    _api_logger = logging.getLogger("api_client")
    _api_client = APIClient(_api_logger, _api_logger)
    try:
        _api_client.login()
    except Exception as e:
        print(f"  API login failed: {e}")
        return

    downloaded = None
    try:
        force_kill_excel()
        downloaded = api_download(company_name)
        print(f"  downloaded: {downloaded}")
        result = process_annual_excel_file(
            downloaded, company_name,
            src_month, src_yy, tgt_month, tgt_yy,
        )
        if result and os.path.exists(result):
            print(f"  SAVED (NOT uploaded): {result}")
            print(f"  --> Open it and confirm the target column has real "
                  f"numbers, not #NAME?.")
        else:
            print("  process returned None (aborted on #NAME? guard or data "
                  "not live). File NOT saved — check the log above.")
    except Exception as e:
        print(f"  error: {e}")
        traceback.print_exc()
    finally:
        force_kill_excel()
        _safe_delete(downloaded)
        if _api_client:
            _api_client.close()


if __name__ == "__main__":
    if "--diagnose" in sys.argv:
        diagnose_ace()
        input("\nPress Enter to exit...")
        sys.exit(0)
    if "--test-refresh" in sys.argv:
        test_refresh()
        input("\nPress Enter to exit...")
        sys.exit(0)
    if "--test-one" in sys.argv:
        idx = sys.argv.index("--test-one")
        name   = sys.argv[idx + 1] if len(sys.argv) > idx + 1 else ""
        period = sys.argv[idx + 2] if len(sys.argv) > idx + 2 else ""
        if not name:
            print('Usage: python annual_processor.py --test-one "Company Name Ltd" [period]')
        else:
            test_one(name, period)
        input("\nPress Enter to exit...")
        sys.exit(0)
    run_forever()
