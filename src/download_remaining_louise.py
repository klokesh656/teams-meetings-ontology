"""
Download remaining Louise recordings from OneDrive with fresh authentication.
Uses direct Graph API calls to get fresh download URLs.
"""

import os
import json
import asyncio
import aiohttp
import requests
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

TENANT_ID = os.getenv("AZURE_TENANT_ID")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

RECORDINGS_DIR = Path("recordings")
REPORT_FILE = Path("output/louise_checkins_report_20260101_004344.json")
PROGRESS_FILE = Path("louise_download_progress.json")


def get_access_token():
    """Get a fresh access token using client credentials."""
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    response = requests.post(url, data=data)
    if response.status_code == 200:
        return response.json().get("access_token")
    else:
        print(f"❌ Failed to get token: {response.text}")
        return None


def get_fresh_download_url(token: str, user_id: str, file_id: str):
    """Get a fresh download URL for a file."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/items/{file_id}"
    headers = {"Authorization": f"Bearer {token}"}
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return data.get("@microsoft.graph.downloadUrl")
    else:
        print(f"    ❌ Failed to get download URL: {response.status_code}")
        return None


async def download_file(url: str, local_path: Path, session: aiohttp.ClientSession):
    """Download a file with progress."""
    try:
        timeout = aiohttp.ClientTimeout(total=1800)  # 30 min timeout
        async with session.get(url, timeout=timeout) as response:
            if response.status == 200:
                local_path.parent.mkdir(parents=True, exist_ok=True)
                total_size = int(response.headers.get('content-length', 0))
                downloaded = 0
                
                with open(local_path, 'wb') as f:
                    async for chunk in response.content.iter_chunked(65536):
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_size > 0:
                            pct = (downloaded / total_size) * 100
                            print(f"\r    Progress: {pct:.1f}% ({downloaded / 1024 / 1024:.1f} MB)", end="", flush=True)
                
                print()  # New line
                return True
            else:
                print(f"    ❌ HTTP {response.status}")
                return False
    except asyncio.TimeoutError:
        print(f"    ❌ Timeout")
        return False
    except Exception as e:
        print(f"    ❌ Error: {str(e)[:50]}")
        return False


def load_progress():
    """Load download progress."""
    if PROGRESS_FILE.exists():
        with open(PROGRESS_FILE, 'r') as f:
            return json.load(f)
    return {"downloaded": [], "transcribed": [], "failed": []}


def save_progress(progress):
    """Save progress."""
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, indent=2)


async def main():
    print("=" * 70)
    print("DOWNLOAD REMAINING LOUISE RECORDINGS FROM ONEDRIVE")
    print("=" * 70)
    
    # Load report
    if not REPORT_FILE.exists():
        print(f"❌ Report file not found: {REPORT_FILE}")
        return
    
    with open(REPORT_FILE, 'r') as f:
        report = json.load(f)
    
    # Get fresh token
    print("\nGetting fresh access token...")
    token = get_access_token()
    if not token:
        print("❌ Failed to authenticate")
        return
    print("✅ Authenticated")
    
    # Load progress
    progress = load_progress()
    
    # Find recordings to download
    all_meetings = report.get("all_meetings", [])
    to_download = []
    
    for meeting in all_meetings:
        if not meeting.get("can_transcribe"):
            continue
        if not meeting.get("has_onedrive_recording"):
            continue
        
        va_name = meeting.get("va_name", "Unknown")
        date = meeting.get("date", "unknown")
        meeting_key = f"{date}_{va_name}"
        
        # Skip if already downloaded
        if meeting_key in progress.get("downloaded", []):
            continue
        
        # Check if file exists locally
        date_str = date.replace("-", "")
        va_name_safe = va_name.replace(" ", "_")
        filename = f"{date_str}_Integration_Team_Check-in_Louise_x_{va_name_safe}.mp4"
        local_path = RECORDINGS_DIR / filename
        
        if local_path.exists() and local_path.stat().st_size > 1000000:
            progress["downloaded"].append(meeting_key)
            save_progress(progress)
            continue
        
        onedrive_info = meeting.get("onedrive_info", {})
        if onedrive_info:
            to_download.append({
                "date": date,
                "va_name": va_name,
                "meeting_key": meeting_key,
                "user_id": onedrive_info.get("user_id"),
                "file_id": onedrive_info.get("file_id"),
                "size_mb": onedrive_info.get("size_mb", 0),
                "filename": filename,
                "local_path": local_path
            })
    
    print(f"\nFound {len(to_download)} recordings to download")
    
    if not to_download:
        print("✅ All recordings already downloaded!")
        return
    
    # Calculate total size
    total_size = sum(r["size_mb"] for r in to_download)
    print(f"Total size: {total_size:.1f} MB")
    
    # Download
    downloaded_count = 0
    failed_count = 0
    
    async with aiohttp.ClientSession() as session:
        for i, rec in enumerate(to_download, 1):
            print(f"\n[{i}/{len(to_download)}] {rec['date']} - Louise x {rec['va_name']}")
            print(f"    Size: {rec['size_mb']:.1f} MB")
            
            # Get fresh download URL
            print("    Getting download URL...")
            download_url = get_fresh_download_url(token, rec["user_id"], rec["file_id"])
            
            if not download_url:
                print("    ❌ Could not get download URL")
                progress["failed"].append({"key": rec["meeting_key"], "reason": "No download URL"})
                save_progress(progress)
                failed_count += 1
                continue
            
            # Download
            print("    Downloading...")
            success = await download_file(download_url, rec["local_path"], session)
            
            if success:
                # Verify file size
                if rec["local_path"].exists():
                    actual_size = rec["local_path"].stat().st_size / 1024 / 1024
                    if actual_size >= rec["size_mb"] * 0.95:  # Within 5%
                        print(f"    ✅ Downloaded: {rec['filename']}")
                        progress["downloaded"].append(rec["meeting_key"])
                        downloaded_count += 1
                    else:
                        print(f"    ⚠️ Incomplete: {actual_size:.1f} MB / {rec['size_mb']:.1f} MB")
                        rec["local_path"].unlink()  # Delete incomplete file
                        progress["failed"].append({"key": rec["meeting_key"], "reason": "Incomplete download"})
                        failed_count += 1
            else:
                progress["failed"].append({"key": rec["meeting_key"], "reason": "Download failed"})
                failed_count += 1
            
            save_progress(progress)
            
            # Refresh token every 10 downloads
            if i % 10 == 0:
                print("\n    Refreshing token...")
                token = get_access_token()
                if not token:
                    print("    ❌ Token refresh failed")
                    break
    
    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"  Downloaded: {downloaded_count}")
    print(f"  Failed: {failed_count}")
    print(f"  Total downloaded: {len(progress.get('downloaded', []))}")
    print(f"\nProgress saved to: {PROGRESS_FILE}")


if __name__ == "__main__":
    asyncio.run(main())
