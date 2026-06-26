"""
API Client — handles authentication, file download, and file upload
to tickercharcha.com.
"""

import os
import time
import logging
import requests

from config import (
    API_LOGIN_URL,
    API_UPLOAD_URL,
    API_DOWNLOAD_URL,
    API_USERNAME,
    API_PASSWORD,
    UPLOAD_MAX_RETRIES,
    UPLOAD_RETRY_DELAY,
)


class APIClient:
    """Manages authentication and file uploads."""

    def __init__(self, logger: logging.Logger, error_logger: logging.Logger):
        self.logger = logger
        self.error_logger = error_logger
        self.token: str | None = None
        self.session = requests.Session()

    # ── Authentication ───────────────────────────────

    def login(self) -> None:
        """
        Authenticate with the API and store the token.
        Raises on failure.
        """
        self.logger.info(f"Authenticating with API at {API_LOGIN_URL}...")
        try:
            # The API expects JSON after all
            response = self.session.post(
                API_LOGIN_URL,
                json={
                    "email": API_USERNAME,
                    "password": API_PASSWORD,
                },
                timeout=30,
            )
            response.raise_for_status()

            data = response.json()

            # Try common token field names
            self.token = (
                data.get("token")
                or data.get("access_token")
                or data.get("auth_token")
                or data.get("data", {}).get("token")
            )
            
            # Extract uploader info for the foreign key constraint
            user_data = data.get("user") or data.get("data", {}).get("user") or data
            self.uploader_name = str(user_data.get("id") or user_data.get("user_id") or user_data.get("name") or API_USERNAME)

            if not self.token:
                raise ValueError(
                    f"Token not found in login response. Keys: {list(data.keys())}"
                )

            self.session.headers.update({
                "Authorization": f"Bearer {self.token}"
            })
            self.logger.info("Authentication successful. Token acquired.")

        except requests.exceptions.HTTPError as e:
            self.logger.error(f"Authentication failed: {e}")
            if e.response is not None:
                self.logger.error(f"Response body: {e.response.text}")
            raise
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Request failed: {e}")
            raise
        except (ValueError, KeyError) as e:
            self.logger.error(f"Failed to parse auth response: {e}")
            raise

    def _ensure_authenticated(self) -> None:
        """Re-authenticate if token is missing."""
        if not self.token:
            self.login()

    # ── File Download ────────────────────────────────

    def download_file(
        self, company_name: str, save_dir: str, max_retries: int = 3
    ) -> str:
        """
        Download a company Excel file from the API.

        Returns the local file path of the saved .xlsx.
        Raises on unrecoverable failure.
        """
        self._ensure_authenticated()

        for attempt in range(1, max_retries + 1):
            try:
                self.logger.info(
                    f"[{company_name}] Downloading via API "
                    f"(attempt {attempt}/{max_retries})..."
                )

                from urllib.parse import quote
                url = f"{API_DOWNLOAD_URL}/{quote(company_name)}"

                response = self.session.get(
                    url,
                    timeout=120,
                    stream=True,
                )

                # Handle token expiration (401/403) → re-auth and retry
                if response.status_code in (401, 403):
                    self.logger.warning(
                        f"[{company_name}] Token expired (HTTP "
                        f"{response.status_code}). Re-authenticating..."
                    )
                    self.token = None
                    self.login()
                    continue

                response.raise_for_status()

                # Determine filename from Content-Disposition or fallback
                filename = f"{company_name}.xlsx"
                cd = response.headers.get("Content-Disposition", "")
                if "filename=" in cd:
                    import re as _re
                    match = _re.search(
                        r'filename[*]?=["\']?([^"\';]+)', cd
                    )
                    if match:
                        filename = match.group(1).strip()

                os.makedirs(save_dir, exist_ok=True)
                local_path = os.path.join(save_dir, filename)

                with open(local_path, "wb") as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)

                file_size = os.path.getsize(local_path)
                self.logger.info(
                    f"[{company_name}] Download complete: "
                    f"{local_path} ({file_size:,} bytes)"
                )
                return local_path

            except requests.exceptions.HTTPError as e:
                self.logger.warning(
                    f"[{company_name}] Download attempt "
                    f"{attempt} failed: {e}"
                )
                if e.response is not None:
                    self.logger.warning(
                        f"Response body: {e.response.text[:500]}"
                    )

                if attempt == max_retries:
                    raise RuntimeError(
                        f"Download failed for '{company_name}' "
                        f"after {max_retries} attempts"
                    ) from e

                time.sleep(UPLOAD_RETRY_DELAY)

            except requests.exceptions.RequestException as e:
                self.logger.warning(
                    f"[{company_name}] Download request error "
                    f"(attempt {attempt}): {e}"
                )
                if attempt == max_retries:
                    raise RuntimeError(
                        f"Download failed for '{company_name}' "
                        f"after {max_retries} attempts"
                    ) from e

                time.sleep(UPLOAD_RETRY_DELAY)

            except Exception as e:
                self.logger.error(
                    f"[{company_name}] Unexpected download error: {e}"
                )
                raise

    # ── File Upload ──────────────────────────────────

    def upload_file(
        self, file_path: str, company_name: str, accord_code: str
    ) -> dict:
        """
        Upload a generated Excel file to the API.
        
        Returns the API response as a dict.
        Raises on unrecoverable failure.
        """
        self._ensure_authenticated()

        filename = os.path.basename(file_path)

        for attempt in range(1, UPLOAD_MAX_RETRIES + 1):
            try:
                self.logger.info(
                    f"[{company_name}] Uploading '{filename}' (attempt {attempt}/{UPLOAD_MAX_RETRIES})..."
                )

                with open(file_path, "rb") as f:
                    response = self.session.post(
                        API_UPLOAD_URL,
                        files={"file": (filename, f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                        data={
                            "company_name": company_name,
                            "uploader_name": getattr(self, 'uploader_name', API_USERNAME),
                            "upload_type": 0,
                        },timeout=120,
                    )

                # Handle token expiration (401/403) → re-auth and retry
                if response.status_code in (401, 403):
                    self.logger.warning(
                        f"[{company_name}] Token expired (HTTP {response.status_code}). Re-authenticating..."
                    )
                    self.token = None
                    self.login()
                    continue

                response.raise_for_status()
                result = response.json()

                self.logger.info(
                    f"[{company_name}] Upload successful. Response: {result}"
                )
                return result

            except requests.exceptions.HTTPError as e:
                self.logger.warning(
                    f"[{company_name}] Upload attempt {attempt} failed: {e}"
                )
                if e.response is not None:
                    self.logger.warning(f"Response body: {e.response.text}")
                
                if attempt == UPLOAD_MAX_RETRIES:
                    raise RuntimeError(
                        f"Upload failed for '{company_name}' after {UPLOAD_MAX_RETRIES} attempts"
                    ) from e
                
                time.sleep(UPLOAD_RETRY_DELAY)
            
            except requests.exceptions.RequestException as e:

                if attempt == UPLOAD_MAX_RETRIES:
                    raise RuntimeError(
                        f"Upload failed for '{company_name}' after {UPLOAD_MAX_RETRIES} attempts"
                    ) from e

                time.sleep(UPLOAD_RETRY_DELAY)

            except Exception as e:
                self.logger.error(
                    f"[{company_name}] Unexpected upload error: {e}"
                )
                raise

    # ── Cleanup ──────────────────────────────────────

    def close(self) -> None:
        """Close the HTTP session."""
        self.session.close()
        self.logger.debug("HTTP session closed.")


if __name__ == "__main__":
    from logger_setup import setup_logger
    import sys

    # 1. Setup logs (we only care about the console output for this test)
    main_logger, error_logger = setup_logger()
    
    # 2. Print what we're doing
    print("="*50)
    print("TESTING API LOGIN INDEPENDENTLY")
    print(f"Target URL: {API_LOGIN_URL}")
    print(f"Username configured: {API_USERNAME}")
    print("="*50)

    # 3. Create client
    client = APIClient(main_logger, error_logger)
    
    try:
        # 4. Test Login
        client.login()
        print("\n✅ LOGIN SUCCESS!")
        
        # 5. Upload all files in the output folder
        import os
        output_dir = "output"
        
        if not os.path.exists(output_dir):
            print(f"\n⚠️ Output folder '{output_dir}' not found.")
            sys.exit(1)
            
        files = [f for f in os.listdir(output_dir) if f.endswith(".xlsx") and not f.startswith("_working_")]
        
        if not files:
            print(f"\n⚠️ No Excel files found in '{output_dir}'.")
            sys.exit(0)
            
        print(f"\nFound {len(files)} files to upload. Starting batch upload...")
        
        success_count = 0
        fail_count = 0
        
        for filename in files:
            file_path = os.path.join(output_dir, filename)
            company_name = filename.replace(".xlsx", "")
            
            print(f"\nUploading: {company_name}")
            try:
                # The accord_code doesn't matter for the upload request as it only sends company_name
                client.upload_file(file_path, company_name=company_name, accord_code="000000")
                print(f"✅ SUCCESS: {company_name}")
                success_count += 1
            except Exception as upload_e:
                print(f"❌ FAILED: {company_name} - {upload_e}")
                fail_count += 1
                
        print("\n" + "="*50)
        print(f"BATCH UPLOAD COMPLETE")
        print(f"Successful: {success_count}")
        print(f"Failed: {fail_count}")
        print("="*50)

    except Exception as e:
        print(f"\n❌ SCRIPT FAILED: {e}")
        sys.exit(1)
    finally:
        client.close()
