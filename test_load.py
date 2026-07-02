import os
import json
import sys
# mock directories
EXE_DIR = os.getcwd()
BASE_DIR = EXE_DIR
DATA_DIR = os.path.join(BASE_DIR, "data")
TRACKER_FILE = os.path.join(DATA_DIR, "annual_tracker.json")

def _normalize_name(name: str) -> str:
    if name is None:
        return ""
    s = str(name).strip().strip('"').strip("'").strip()
    s = s.rstrip(".")
    s = " ".join(s.split())
    return s.casefold()

def _load_done_today() -> set:
    result = set()
    try:
        if os.path.exists(TRACKER_FILE):
            with open(TRACKER_FILE, "r", encoding="utf-8") as f:
                all_data = json.load(f)
            for date_key, day_data in all_data.items():
                for company_name, entry in day_data.items():
                    if entry.get("status") == "uploaded":
                        period = str(entry.get("period", ""))
                        result.add((_normalize_name(company_name), period))
    except Exception as e:
        print(e)
    return result

done = _load_done_today()
print(f"Loaded {len(done)} companies")
for d in list(done)[:10]:
    print(d)
