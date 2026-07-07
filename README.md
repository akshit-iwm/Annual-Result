# Annual Result Processor v2 (API Mode)

Automated pipeline for extracting, processing, and uploading annual financial results. Runs 24×7 on a Windows machine with Microsoft Excel and the ACE add-in installed. Downloads company templates from the Ticker Charcha API, copies the previous year's formula column into the new year, refreshes ACE data, validates results (PAT & Sources of Funds), and uploads the completed file back.

---

## Table of Contents

- [System Architecture](#system-architecture)
- [Directory Structure](#directory-structure)
- [Execution Workflow](#execution-workflow)
- [Excel Processing — Step by Step](#excel-processing--step-by-step)
- [Deduplication & Skip Logic](#deduplication--skip-logic)
- [Auto-Dismiss Excel Dialogs](#auto-dismiss-excel-dialogs)
- [Environment Variables](#environment-variables)
- [Installation & Setup](#installation--setup)
- [Usage](#usage)
- [Troubleshooting](#troubleshooting)

---

## System Architecture

### Core Modules

| File | Purpose |
|------|---------|
| `annual_processor.py` | Main orchestrator — multi-threaded (Poller + Worker), 24×7 resilient loop, Excel COM automation, formula copying, ACE refresh, PAT/SOF validation, API upload, Slack notifications |
| `api_client.py` | Handles authentication (Bearer token), file download, and file upload to the Ticker Charcha API. Auto-retries on 401/403 with re-authentication |
| `config.py` | Loads secrets and settings from `.env` via `python-dotenv` (API URLs, credentials, file paths, retry counts) |
| `diagnose_ace_addin.py` | Diagnostic utility to inspect the ACE Excel add-in installation — lists COM objects, ribbon XML, and exposed macro callbacks |

### Threading Model

```
┌─────────────────────────────┐
│        run_forever()        │  ← 24×7 resilience wrapper (auto-restart on crash)
│                             │
│  ┌───────────────────────┐  │
│  │  ACE Annual Poller    │  │  ← Reads ACE Excel every 30 min, queues new companies
│  │  (daemon thread)      │  │
│  └────────┬──────────────┘  │
│           │ PriorityQueue   │
│  ┌────────▼──────────────┐  │
│  │  Annual Worker        │  │  ← Picks companies, downloads, processes Excel, uploads
│  │  (daemon thread)      │  │
│  └───────────────────────┘  │
│                             │
│  ┌───────────────────────┐  │
│  │  Dialog Dismisser     │  │  ← Auto-clicks OK on Excel popup dialogs (Circular Ref, etc.)
│  │  (daemon thread)      │  │
│  └───────────────────────┘  │
│                             │
│  Main thread: status logs,  │
│  24-hour API token refresh  │
└─────────────────────────────┘
```

---

## Directory Structure

```
Annual-Result/
├── annual_processor.py        # Main script
├── api_client.py              # API client (login, download, upload)
├── config.py                  # .env config loader
├── diagnose_ace_addin.py      # ACE add-in diagnostic tool
├── .env                       # Secrets (API URLs, credentials) — NOT committed
├── Annual Update Ace.xlsx     # Master ACE workbook (company list source)
├── CharchaCompaniesList.txt   # Allow-list of valid Charcha company names
├── requirements.txt           # Python dependencies
│
├── result_files/              # Temp: downloaded raw Excel files (auto-cleaned)
├── Updated_Excel/             # Output: processed & uploaded Excel files
├── Archived/                  # Historical metadata backups
│
├── data/
│   ├── annual_tracker.json    # Per-day processing status for every company
│   └── metadata.json          # Legacy upload log (company → timestamp)
│
├── annual_processor.log       # Full runtime log
├── skipped_log_annual.txt     # Append-only log of skipped/failed companies
└── logs.txt                   # Legacy log file
```

---

## Execution Workflow

### 1. Polling (`ACE Annual Poller` thread)

- Opens `Annual Update Ace.xlsx` in a dedicated Excel instance every 30 minutes
- Loads the ACE add-in (both `.xll` UDF provider and COM data plumbing)
- Triggers the ACE ribbon Refresh button via Win32 keyboard simulation (`Alt → Y2 → Y8`)
- Reads company names (column B) and P&L periods (column I) from row 3 onward
- For each company:
  - Checks the **Charcha allow-list** (`CharchaCompaniesList.txt`) — skips if not found
  - Checks if the file already exists in `Updated_Excel/` — skips if found
  - Checks in-memory `done_today` and `failed_today` sets — skips if already processed or failed
  - Queues new companies into a `PriorityQueue` with a configurable initial delay

### 2. Processing (`Annual Worker` thread)

- Picks the next company from the queue (respects scheduled time)
- Downloads the company's Excel template via the API
- Opens it using `os.startfile` (preserves ACE add-in loading)
- Performs the [10-step Excel processing](#excel-processing--step-by-step)
- If successful: saves to `Updated_Excel/`, uploads via API, sends Slack notification
- If data not yet live (PAT unchanged): deletes file, re-queues with a retry delay
- If error (download/Excel/upload failure): marks failed, moves to next company

### 3. Resilience

- `run_forever()` wrapper catches unhandled exceptions, logs traceback, waits 60s, restarts
- API token is proactively refreshed every 24 hours
- Midnight date change triggers a full state reset (clears `failed_today`, reloads `done_today`)

---

## Excel Processing — Step by Step

This is what happens inside `process_annual_excel_file()` for each company:

| Step | Description |
|------|-------------|
| **1. Scan markers** | Reads Column A (rows 1–500) to find section boundaries: P&L Start, Valuation End, Presentation Data Start/End, Balance Sheet Start/End, PAT row, Sources of Funds row |
| **2. Find QTR marker** | Scans header row 1 for `QTR` or `QUARTERLY` — this marks the right boundary of the annual section |
| **3. Find source & target columns** | In the annual section (columns before QTR marker), finds the source column (e.g. `Mar-25`) and target column (e.g. `Mar-26`) by matching dates. Falls back to the two most recent dated columns if the ACE period doesn't match (fiscal year change) |
| **4. Build copy ranges** | Creates row ranges from P&L Start to last used row, **excluding Presentation Data rows** |
| **5. Write target header** | If the target column doesn't exist yet, creates it with the correct date header |
| **5b. Backup to Historical Estimates** | Copies existing target column values into a "Historical Estimates" column. Uses **first-write-wins** policy — never overwrites an existing backup |
| **6. Copy formulas** | Copies the source column's formulas → target column using `Range.Copy(Destination=...)`. This native Excel copy operation automatically updates relative cell references in formulas, making it mathematically identical to dragging (AutoFill) the formula column in Excel. |
| **7. Pre-refresh PAT** | Records the PAT value before ACE refresh |
| **8. ACE data refresh** | Triggers ACE ribbon Refresh via Win32 keystrokes, then calls `RefreshAll` + `CalculateFull` |
| **9. Post-refresh validation** | Compares pre vs post PAT. Checks Sources of Funds > 0. Checks for `#NAME?` errors (ACE add-in failure — aborts save to prevent corruption) |
| **10. Save** | Saves to `Updated_Excel/{Company Name}.xlsx` |

---

## Deduplication & Skip Logic

The system uses **multiple layers** to prevent re-processing companies:

1. **`Updated_Excel/` folder check** (filesystem) — if `{Company Name}.xlsx` exists, skip instantly
2. **`done_today` set** (in-memory) — tracks `(normalized_name, period)` tuples of successfully uploaded companies across all dates from `annual_tracker.json`
3. **`failed_today` set** (in-memory) — companies that errored today are not retried until the next day
4. **`in_queue` set** (in-memory) — prevents duplicate queue entries within the same poll cycle
5. **Charcha allow-list** (`CharchaCompaniesList.txt`) — companies not in this list are skipped without attempting a download
6. **`metadata.json`** (legacy) — backward compatibility with older runs

Company names are **normalized** (case-folded, whitespace-collapsed, quotes/dots stripped) before comparison, so `"ABB India Ltd."` and `"abb india ltd"` are treated as the same company.

---

## Auto-Dismiss Excel Dialogs

When Excel files are opened via `os.startfile`, modal dialogs (e.g. **Circular Reference** warnings, **Update Links** prompts) appear **before** the script can set `display_alerts = False`. These dialogs block the entire automation.

The **Dialog Dismisser** thread solves this:
- Runs in the background, scanning every 1.5 seconds for visible windows titled "Microsoft Excel"
- Uses Win32 API (`EnumWindows`, `EnumChildWindows`, `PostMessage`) to find and click the "OK" or "Yes" button
- Starts before `os.startfile` is called, stops after the workbook is closed
- Logs every dismissed dialog for auditability

---

## Environment Variables

### In `.env` file (loaded by `config.py`)

| Variable | Description |
|----------|-------------|
| `API_LOGIN_URL` | Ticker Charcha authentication endpoint |
| `API_UPLOAD_URL` | File upload endpoint |
| `API_DOWNLOAD_URL` | File download endpoint (company name appended as URL path) |
| `API_USERNAME` | Login email |
| `API_PASSWORD` | Login password |
| `UPLOAD_MAX_RETRIES` | Max upload retry attempts (default: 3) |
| `UPLOAD_RETRY_DELAY` | Seconds between upload retries (default: 5) |

### Runtime overrides (set in OS environment or shell)

| Variable | Default | Description |
|----------|---------|-------------|
| `ANNUAL_POLL_INTERVAL` | `1800` | Seconds between ACE Excel polls (30 min) |
| `ANNUAL_INITIAL_DELAY_MINS` | `1` | Minutes to wait before processing a newly queued company |
| `ANNUAL_RETRY_DELAY_MINS` | `1` | Minutes before retrying a company whose data isn't live yet |
| `EXCEL_REFRESH_WAIT` | `20` | Seconds to wait after ACE data refresh |
| `ACE_REFRESH_WAIT` | `300` | Max seconds to wait for ACE master refresh to complete |
| `EXCEL_OPEN_WAIT` | `6` | Seconds to wait after `os.startfile` before attempting to attach |
| `ACE_XLL_PATH` | `E:\Ace Equity Nxt\ActiveXl\ACEEQ_XL_NXT64.xll` | Path to the ACE Excel add-in |
| `ACE_COM_PROGID` | `ACEEQ_XL_NXT` | COM ProgId for the ACE add-in |
| `ACE_RIBBON_TAB_KEYS` | `Y2` | KeyTip for the ACE ribbon tab |
| `ACE_RIBBON_REFRESH_KEY` | `Y8` | KeyTip for the ACE Refresh button |
| `RUN_LOOP_COOLDOWN` | `60` | Seconds to wait before restarting after an unhandled crash |

---

## Installation & Setup

### Prerequisites

- **Windows** with Microsoft Excel installed (COM automation required)
- **ACE Equity Nxt** add-in installed and licensed
- **Python 3.10+**

### Steps

1. **Clone the repository:**
   ```bash
   git clone https://github.com/akshit-iwm/Annual-Result.git
   cd Annual-Result
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Create `.env` file** in the project root:
   ```env
   API_LOGIN_URL=https://api.tickercharcha.com/api/login
   API_UPLOAD_URL=https://api.tickercharcha.com/api/upload-excel
   API_DOWNLOAD_URL=https://api.tickercharcha.com/api/download
   API_USERNAME=your_email@example.com
   API_PASSWORD=your_password
   ```

4. **Place required files:**
   - `Annual Update Ace.xlsx` — the ACE master workbook (in the project root)
   - `CharchaCompaniesList.txt` — one company name per line (in the project root)

5. **Verify ACE add-in** loads correctly:
   ```bash
   python annual_processor.py --diagnose
   ```

---

## Usage

### 24×7 Production Mode

```bash
python annual_processor.py
```

Runs continuously. Press `Ctrl+C` for clean shutdown.

### Test Single Company (no upload)

Process one company end-to-end, save to `Updated_Excel/` but do NOT upload. Useful for manual inspection:

```bash
python annual_processor.py --test-one "Company Name Ltd"
python annual_processor.py --test-one "Castrol India Ltd" 202512
```

The optional second argument is the ACE period code (e.g. `202512` = Dec-25 → Dec-25).

### Test ACE Ribbon Refresh

Isolated test of the ACE ribbon refresh to verify keyboard simulation works:

```bash
python annual_processor.py --test-refresh
```

### Diagnose ACE Add-In

Inspect Excel COM objects, add-in registration, and ribbon XML:

```bash
python annual_processor.py --diagnose
```

---

## Troubleshooting

| Problem | Cause | Fix |
|---------|-------|-----|
| Script hangs on "Circular reference" dialog | Excel popup blocks before `display_alerts=False` is set | Fixed: auto-dismiss thread clicks OK automatically |
| `#NAME?` errors in output | ACE add-in not loaded in Excel instance | Use `os.startfile` (not `xw.App`). Run `--diagnose` to verify add-in |
| Same companies processed repeatedly | Dedup was matching by name only, not by period | Fixed: now tracks `(name, period)` tuples + checks `Updated_Excel/` folder |
| API token expired | 48-hour token lifetime | Fixed: proactive 24-hour refresh cycle |
| ACE company list not growing | Ribbon Refresh not triggered | Verify `ACE_RIBBON_TAB_KEYS` and `ACE_RIBBON_REFRESH_KEY` with `--test-refresh` |
| Excel zombie processes | Script crashed mid-processing | `force_kill_excel()` runs before each company + on errors |

### Log Files

- **`annual_processor.log`** — full runtime log (all operations, timestamps, errors)
- **`skipped_log_annual.txt`** — append-only record of every skipped/failed company with reason
- **`data/annual_tracker.json`** — structured JSON tracker, organized by date, with status per company

### Slack Notifications

Successfully processed companies trigger a Slack message via the configured webhook. The message format is:
```
{Company Name} annual result updated on DB
```
