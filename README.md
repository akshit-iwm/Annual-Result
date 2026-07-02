# Annual Result Processor v2 (API Mode)

This pipeline automates the extraction, processing, and validation of annual financial results. Operating 24x7 in a highly resilient architecture, it handles file downloads, interacts with an Excel COM-based pipeline utilizing the ACE Excel add-in, performs critical data integrity checks, and directly uploads processed files via the Ticker Charcha API.

## System Architecture

### Core Modules
* **`annual_processor.py`**: The main orchestrator of the pipeline. It features a multi-threaded architecture (Poller and Worker threads) that runs continuously. It coordinates downloading base templates, manipulating Excel objects to copy formulas, programmatically refreshing the ACE data plugin, validating results, and uploading completed files. It handles exceptions dynamically and restarts via a `run_forever()` wrapper.
* **`api_client.py`**: Dedicated API handler responsible for authentication, file downloads, and file uploads to the Ticker Charcha API.
* **`config.py`**: Configuration parser leveraging `python-dotenv` to load secrets and operational parameters (e.g., API URLs, file paths).
* **`diagnose_ace_addin.py`**: A utility script used to inspect, debug, and trace the ACE Excel add-in installation, identifying the specific COM objects and macros needed for programmatic interaction.

### Execution Workflow
1. **Polling (`ACE Annual Poller`)**: A background thread polls the designated master Excel file periodically to determine which companies are due for an annual result update. It checks against a Charcha company allow-list before queuing them.
2. **Downloading**: The worker thread picks up queued companies, validates them, and invokes the API to download the latest raw Excel data.
3. **Excel Processing (COM Automation)**:
   - Uses `xlwings`/`win32com` to open the Excel workbook.
   - Copies financial formula blocks from the previous year (e.g., Mar-25) to the target year (e.g., Mar-26).
   - Simulates a ribbon click/macro execution to force the ACE add-in to refresh data from its backend.
   - Performs validation: It compares pre- and post-refresh PAT (Profit After Tax) values and validates the Source of Funds. Crucially, it aborts saving if the target column is flooded with `#NAME?` errors (which would indicate the ACE plugin failed to load).
4. **Uploading**: Successfully refreshed and validated Excel workbooks are saved and uploaded via the API.
5. **Persistence & Resilience**: 
   - Proactive 24-hour API token refreshes prevent session expiration.
   - Any fatal crashes in the main logic invoke an automatic cooldown and loop restart.

## Installation & Setup

1. **Clone & Environment Setup**:
   Ensure you are using Python 3.x and have a virtual environment configured.
   ```bash
   pip install -r requirements.txt
   ```

2. **Environment Variables**:
   Create a `.env` file in the root directory based on the expected values in `config.py`:
   ```env
   # API Configuration
   API_LOGIN_URL="https://..."
   API_UPLOAD_URL="https://..."
   API_DOWNLOAD_URL="https://..."
   API_USERNAME="your_username"
   API_PASSWORD="your_password"

   # File paths
   CSV_FILE_PATH="data/list.csv"
   TEMPLATE_FILE_PATH="data/template.xlsx"
   OUTPUT_FOLDER_PATH="Updated_Excel"
   
   # Excel and App Configuration
   REFRESH_WAIT_SECONDS=30
   EXCEL_MAX_RETRIES=3
   UPLOAD_MAX_RETRIES=3
   ```

3. **Prerequisites (Windows)**:
   This application heavily relies on Excel COM interactions and the ACE Add-in. It must be run on a Windows machine where Microsoft Excel and the ACE Add-in are properly installed and authorized.

## Usage

### 24x7 Production Mode
To launch the continuous processing loop:
```bash
python annual_processor.py
```
This mode manages its own API session lifecycle, Excel process killings (to avoid zombie processes), and queue handling.

### Diagnostic & Testing Tools
The script includes embedded sub-commands for localized debugging and verifying changes without impacting the production database:

* **Single Company Test**
  Test the full download -> process -> save pipeline for a single company without uploading it back. Useful for manual inspection of output.
  ```bash
  python annual_processor.py --test-one "Company Name Ltd" [period]
  ```

* **Refresh Testing**
  Runs an isolated test of the ACE ribbon refresh macro logic to ensure Excel communicates properly with the ACE database.
  ```bash
  python annual_processor.py --test-refresh
  ```

* **Diagnose ACE Add-In**
  Scans the Excel application instance and the ACE installation directory to identify COM objects, Ribbon XML, and exposed macro functions.
  ```bash
  python annual_processor.py --diagnose
  ```
