"""Quick test: login + download a company Excel via API."""
import os
import sys

# Ensure we can import config / api_client from the same directory
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from api_client import APIClient
import logging

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("test")

SAVE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "result_files")

company = "Reliance Industries Ltd"
if len(sys.argv) > 1:
    company = sys.argv[1]

print("=" * 50)
print(f"TEST DOWNLOAD: '{company}'")
print(f"Save dir: {SAVE_DIR}")
print("=" * 50)

client = APIClient(logger, logger)
try:
    client.login()
    print("\n✅ Login OK\n")

    path = client.download_file(company, SAVE_DIR)
    size = os.path.getsize(path)
    print(f"\n✅ Download OK: {path}  ({size:,} bytes)")
except Exception as e:
    print(f"\n❌ FAILED: {e}")
    import traceback; traceback.print_exc()
finally:
    client.close()
